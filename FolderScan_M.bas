Attribute VB_Name = "FolderScan_M"
Option Explicit

Public Sub FolderScan()
    Dim ws As Worksheet
    Dim loConfig As ListObject, loFiles As ListObject
    Dim found As Range
    Dim aLocalRoot As String, domainName As String, htmlFilePath As String
    Dim urlPrefix As String, urlSuffix As String

    Dim r As ListRow
    Dim existingRows As Object          ' Scripting.Dictionary
    Dim existingFolderSet As Object     ' Scripting.Dictionary
    Dim showSettings As Object          ' Scripting.Dictionary  (only from rows present in Excel)
    Dim allDescs As Collection
    Dim desc As FileDescriptor

    Dim objType As String, showVal As String
    Dim cellVal As Variant
    Dim maxNum As Long, nextNum As Long

    Dim colNumber As Long, colDateFound As Long, colKeywords As Long
    Dim colDomain As Long, colCategory As Long, colFolder As Long
    Dim colObjType As Long, colFileName As Long, colObjName As Long
    Dim colRelPath As Long, colShow As Long, colLink As Long

    Dim rel As Variant, relKey As String
    Dim skip As Boolean
    Dim parts() As String, n As Long
    Dim isFolderObj As Boolean

    Dim highlightRanges As Collection
    Dim hl As Range, rngUnion As Range

    Dim allowed As Object               ' RelativePath -> FileDescriptor (after Show? rules)
    Dim baseNewSet As Object            ' RelativePath -> True (new folders to base-limit)
    Dim showDefaults As Object          ' RelativePath -> "1st Level"/"Nothing"/""
    Dim keysArr As Variant, kKey As Variant, childKey As Variant
    Dim toRemove As Collection
    Dim parentPath As String, p As Long, diff As Long

    Dim remainingNew As Object
    Dim newKeys As Variant, numRows As Long, numCols As Long
    Dim batchData() As Variant, i As Long
    Dim startRow As Long, targetRange As Range
    Dim rowToUpdate As ListRow, c As Range

    Dim fullPath As String, t As String
    Dim rp As String
    Dim effRule As String, effRulePath As String
    Dim tmpPath As String

    ' Numbering helpers
    Dim blankNumberCells As Collection
    Dim isYellow As Boolean

    ' Link build vars (used only for NEW rows)
    Dim encRel As String, computedLink As String, isShortcut As Boolean

    ' New rows fill color (Light Blue)
    Const NEWROW_COLOR As Long = &HF7EBDD   ' RGB(221,235,247)

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    '— 1) tbConfig: Local Root, Domain name, Url Prefix / Suffix —
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

    ' Url Prefix / Url Suffix (blank if not present)
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

    '— 2) Locate tbFiles —
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Set loFiles = ws.ListObjects("tbFiles")
        On Error GoTo 0
        If Not loFiles Is Nothing Then Exit For
    Next ws
    If loFiles Is Nothing Then MsgBox "Output table 'tbFiles' not found.", vbCritical: GoTo CleanFail

    ' Cache column indexes
    colNumber = loFiles.ListColumns("#").Index
    colDateFound = loFiles.ListColumns("Date found").Index
    colKeywords = loFiles.ListColumns("Keywords").Index
    colDomain = loFiles.ListColumns("Domain").Index
    colCategory = loFiles.ListColumns("Category").Index
    colFolder = loFiles.ListColumns("Folder").Index
    colObjType = loFiles.ListColumns("Object Type").Index
    colFileName = loFiles.ListColumns("Filename").Index
    colObjName = loFiles.ListColumns("Object name").Index
    colRelPath = loFiles.ListColumns("RelativePath").Index
    colLink = IIf(ColumnExists(loFiles, "Link"), loFiles.ListColumns("Link").Index, 0)
    colShow = IIf(ColumnExists(loFiles, "Show?"), loFiles.ListColumns("Show?").Index, 0)

    ' === Clear all existing background before scanning (reset yellow/blue) ===
    If Not loFiles.DataBodyRange Is Nothing Then
        loFiles.DataBodyRange.Interior.ColorIndex = xlColorIndexNone
    End If
    ' ========================================================================

    '— 3) Read existing rows, folders & Show? rules; track max "#", collect blanks (ignoring yellow rows) —
    Set existingRows = CreateObject("Scripting.Dictionary")
    Set existingFolderSet = CreateObject("Scripting.Dictionary")
    Set showSettings = CreateObject("Scripting.Dictionary")
    Set blankNumberCells = New Collection
    maxNum = 0

    For Each r In loFiles.ListRows
        ' Keep existing rows dictionary (used by show rules, etc.)
        If Not IsError(r.Range.Cells(1, colRelPath).Value) Then
            relKey = CStr(r.Range.Cells(1, colRelPath).Value)
            If existingRows.Exists(relKey) Then existingRows.Remove relKey
            existingRows.Add relKey, r

            objType = CStr(r.Range.Cells(1, colObjType).Value)
            If objType = "Category" Or objType = "Folder" Or objType = "Subfolder" Then
                If Not existingFolderSet.Exists(relKey) Then existingFolderSet.Add relKey, True
            End If
        End If

        ' Determine if this row is yellow (missing on disk) — ignore for numbering
        isYellow = False
        On Error Resume Next
        isYellow = (r.Range.Cells(1, 1).Interior.Color = vbYellow) Or (r.Range.Cells(1, 1).Interior.ColorIndex = 6)
        On Error GoTo 0

        ' Collect Show? from folders (yellow rows can still carry rules)
        objType = CStr(r.Range.Cells(1, colObjType).Value)
        If objType = "Category" Or objType = "Folder" Or objType = "Subfolder" Then
            If colShow > 0 Then
                showVal = LCase$(Trim$(CStr(r.Range.Cells(1, colShow).Value)))
                If Len(showVal) > 0 Then
                    If showSettings.Exists(relKey) Then showSettings.Remove relKey
                    showSettings.Add relKey, showVal
                End If
            End If
        End If

        ' Max-number and blank-number handling (ignore yellow rows)
        If Not isYellow Then
            cellVal = r.Range.Cells(1, colNumber).Value
            If Len(Trim$(CStr(cellVal))) = 0 Then
                blankNumberCells.Add r.Range.Cells(1, colNumber)
            ElseIf IsNumeric(cellVal) Then
                If CLng(cellVal) > maxNum Then maxNum = CLng(cellVal)
            End If
        End If
    Next r
    nextNum = maxNum + 1

    '— 4) Scan filesystem —
    Set allDescs = New Collection
    ScanFoldersCollectDescriptors aLocalRoot, allDescs

    '— 4b) Map fields; relabel folders (L1=Category, L2=Folder, L3+=Subfolder) and restore shortcut link extraction —
    For Each desc In allDescs
        desc.Domain = domainName

        isFolderObj = isFolder(desc.ObjectType)
        parts = Split(desc.RelativePath, "\")
        n = UBound(parts)

        If isFolderObj Then
            If n >= 0 Then desc.Category = parts(0) Else desc.Category = ""
            If n >= 1 Then desc.Folder = parts(1) Else desc.Folder = ""

            Select Case n
                Case 0: desc.ObjectType = "Category"
                Case 1: desc.ObjectType = "Folder"
                Case Else: desc.ObjectType = "Subfolder"
            End Select
        Else
            If n >= 1 Then desc.Category = parts(0) Else desc.Category = ""
            If n >= 2 Then desc.Folder = parts(1) Else desc.Folder = ""

            ' SHORTCUTS: read .lnk / .url into desc.Link (TargetPath / URL=)
            If EndsWithCI(desc.fileName, ".lnk") Or EndsWithCI(desc.fileName, ".url") Then
                fullPath = aLocalRoot & desc.RelativePath
                desc.Link = ResolveShortcutLink(fullPath)   ' may be ""
                t = DeriveShortcutTypeFromName(desc.fileName)
                If Len(t) > 0 Then
                    desc.ObjectType = LCase$(t)
                Else
                    desc.ObjectType = "shortcut"
                End If
            End If
        End If
    Next desc

    '— 5) Apply Show? rules with INNER LEVEL OVERRIDE (nearest ancestor in Excel wins) —
    Set allowed = CreateObject("Scripting.Dictionary")

    For Each desc In allDescs
        rp = desc.RelativePath
        effRule = ""
        effRulePath = ""

        tmpPath = rp
        Do While Len(tmpPath) > 0
            If showSettings.Exists(tmpPath) Then
                effRule = CStr(showSettings(tmpPath))
                effRulePath = tmpPath
                Exit Do
            End If
            p = InStrRev(tmpPath, "\")
            If p > 0 Then
                tmpPath = Left$(tmpPath, p - 1)
            Else
                Exit Do
            End If
        Loop

        skip = False
        If Len(effRule) > 0 Then
            Select Case effRule
                Case "all"
                    ' keep everything
                Case "nothing"
                    If LCase$(rp) <> LCase$(effRulePath) Then skip = True
                Case "subfolders"
                    If LCase$(rp) <> LCase$(effRulePath) Then
                        If Not (desc.ObjectType = "Category" Or desc.ObjectType = "Folder" Or desc.ObjectType = "Subfolder") Then
                            skip = True
                        End If
                    End If
                Case "1st level"
                    If LCase$(rp) <> LCase$(effRulePath) Then
                        diff = UBound(Split(rp, "\")) - UBound(Split(effRulePath, "\"))
                        If diff > 1 Then skip = True
                    End If
                Case Else
                    ' treat unknown as "all"
            End Select
        End If

        If Not skip Then
            relKey = CStr(desc.RelativePath)
            If allowed.Exists(relKey) Then allowed.Remove relKey
            allowed.Add relKey, desc
        End If
    Next desc

    '— 6) New-folder policy (defaults for folders not present in Excel) —
    Set baseNewSet = CreateObject("Scripting.Dictionary")
    Set showDefaults = CreateObject("Scripting.Dictionary")

    keysArr = allowed.keys

    ' 6a) Identify base new folders
    For Each kKey In keysArr
        relKey = CStr(kKey)
        Set desc = allowed(relKey)
        If desc.ObjectType = "Category" Or desc.ObjectType = "Folder" Or desc.ObjectType = "Subfolder" Then
            If Not existingRows.Exists(relKey) Then
                p = InStrRev(relKey, "\")
                If p > 0 Then
                    parentPath = Left$(relKey, p - 1)
                Else
                    parentPath = ""
                End If
                If (parentPath = "" Or existingFolderSet.Exists(parentPath)) Then
                    If Not baseNewSet.Exists(relKey) Then baseNewSet.Add relKey, True
                End If
            End If
        End If
    Next kKey

    ' 6b) For each base new folder: set Show? default based on PARENT'S explicit rule
    '     - If parent has Show? = "1st Level" (explicit in Excel) => this new folder => "Nothing"
    '     - Else => this new folder => "1st Level"
    Set toRemove = New Collection
    For Each kKey In baseNewSet.keys
        relKey = CStr(kKey)

        ' Determine explicit parent rule (if any)
        p = InStrRev(relKey, "\")
        If p > 0 Then
            parentPath = Left$(relKey, p - 1)
        Else
            parentPath = ""
        End If

        Dim parentRule As String
        parentRule = ""
        If Len(parentPath) > 0 Then
            If showSettings.Exists(parentPath) Then
                parentRule = LCase$(CStr(showSettings(parentPath)))
            End If
        End If

        ' Set default for the new folder itself
        If parentRule = "1st level" Then
            If Not showDefaults.Exists(relKey) Then showDefaults.Add relKey, "Nothing"
        Else
            If Not showDefaults.Exists(relKey) Then showDefaults.Add relKey, "1st Level"
        End If

        ' Trim deeper descendants in this run (same behavior as before)
        For Each childKey In keysArr
            rp = CStr(childKey)
            If LCase$(Left$(rp, Len(relKey) + 1)) = LCase$(relKey & "\") Then
                diff = UBound(Split(rp, "\")) - UBound(Split(relKey, "\"))
                Set desc = allowed(rp)
                If diff > 1 Then
                    toRemove.Add rp
                End If
            End If
        Next childKey
    Next kKey

    For i = 1 To toRemove.Count
        If allowed.Exists(toRemove(i)) Then allowed.Remove toRemove(i)
    Next i

    '— 7) Highlight missing; update existing (NO Link changes); gather remaining new —
    Set highlightRanges = New Collection
    Set remainingNew = CreateObject("Scripting.Dictionary")

    For Each rel In existingRows.keys
        relKey = CStr(rel)
        If Not allowed.Exists(relKey) Then
            highlightRanges.Add existingRows(relKey).Range
        Else
            Set rowToUpdate = existingRows(relKey)
            Set desc = allowed(relKey)
            With rowToUpdate.Range
                .Cells(1, colDomain).Value = desc.Domain
                .Cells(1, colCategory).Value = desc.Category
                .Cells(1, colFolder).Value = desc.Folder
                .Cells(1, colObjType).Value = desc.ObjectType
                .Cells(1, colFileName).Value = desc.fileName
                .Cells(1, colObjName).Value = desc.ObjectName
                .Cells(1, colRelPath).Value = desc.RelativePath
                ' IMPORTANT: Do NOT modify Link for existing rows (leave as-is even if blank)
                .Cells(1, colKeywords).Value = CleanKeywords(.Cells(1, colKeywords).Value)
                .Interior.ColorIndex = xlColorIndexNone
            End With
            allowed.Remove relKey
        End If
    Next rel

    If highlightRanges.Count > 0 Then
        For Each hl In highlightRanges
            If rngUnion Is Nothing Then
                Set rngUnion = hl
            Else
                Set rngUnion = Application.Union(rngUnion, hl)
            End If
        Next hl
        If Not rngUnion Is Nothing Then rngUnion.Interior.Color = vbYellow
    End If

    For Each rel In allowed.keys
        remainingNew.Add CStr(rel), allowed(rel)
    Next rel

    ' ===== FILL BLANK "#" IN EXISTING (NON-YELLOW) ROWS BEFORE ADDING NEW ROWS =====
    If blankNumberCells.Count > 0 Then
        For Each c In blankNumberCells
            If Len(Trim$(CStr(c.Value))) = 0 Then
                c.Value = nextNum
                nextNum = nextNum + 1
            End If
        Next c
    End If
    ' ==============================================================================

    '— 8) Batch add new rows (sorted), assigning Show? defaults —
    newKeys = remainingNew.keys
    If Not IsEmpty(newKeys) Then
        If UBound(newKeys) > LBound(newKeys) Then QuickSortArray newKeys, LBound(newKeys), UBound(newKeys)
    End If

    If remainingNew.Count > 0 Then
        numRows = remainingNew.Count
        numCols = loFiles.ListColumns.Count
        ReDim batchData(1 To numRows, 1 To numCols)

        i = 1
        For Each rel In newKeys
            relKey = CStr(rel)
            Set desc = remainingNew(relKey)

            ' Decide shortcut vs non-shortcut
            isShortcut = EndsWithCI(desc.fileName, ".lnk") Or EndsWithCI(desc.fileName, ".url")

            ' Build link for NEW rows only:
            '   - Shortcuts: desc.Link (extracted)
            '   - Others:    computed prefix/encodedPath/suffix
            If isShortcut Then
                computedLink = Trim$(desc.Link)
            Else
                encRel = Replace(desc.RelativePath, "\", "/")
                encRel = Replace(encRel, "%", "%25")  ' encode % first
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

            batchData(i, colDomain) = desc.Domain
            batchData(i, colCategory) = desc.Category
            batchData(i, colFolder) = desc.Folder
            batchData(i, colObjType) = desc.ObjectType
            batchData(i, colFileName) = desc.fileName
            batchData(i, colObjName) = desc.ObjectName
            batchData(i, colRelPath) = desc.RelativePath
            If colLink > 0 Then batchData(i, colLink) = computedLink
            batchData(i, colNumber) = nextNum
            batchData(i, colDateFound) = Date
            batchData(i, colKeywords) = ""  ' normalize after write

            ' Assign Show? default captured earlier for this new item (if any)
            If colShow > 0 Then
                If showDefaults.Exists(relKey) Then
                    batchData(i, colShow) = showDefaults(relKey)  ' "1st Level" or "Nothing"
                Else
                    batchData(i, colShow) = ""                    ' leave blank (files, etc.)
                End If
            End If

            nextNum = nextNum + 1
            i = i + 1
        Next rel

        ' add rows then dump in one go
        For i = 1 To numRows
            loFiles.ListRows.Add
        Next i

        startRow = loFiles.DataBodyRange.Rows.Count - numRows + 1
        Set targetRange = loFiles.DataBodyRange.Rows(startRow).Resize(numRows, numCols)
        targetRange.Value = batchData

        ' Paint NEW rows Light Blue
        targetRange.Interior.Color = NEWROW_COLOR

        ' clean keywords for new rows
        If colKeywords > 0 Then
            For Each c In targetRange.Columns(colKeywords).Cells
                c.Value = CleanKeywords(c.Value)
            Next c
        End If
    End If

    '— 9) Make all Link cells clickable —
    If colLink > 0 And Not loFiles.DataBodyRange Is Nothing Then
        ApplyHyperlinksOnLinkColumn loFiles, colLink
    End If

    '— 10) Sort tbFiles by RelativePath —
    With loFiles.Sort
        .SortFields.Clear
        .SortFields.Add key:=loFiles.ListColumns("RelativePath").DataBodyRange, _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .Apply
    End With

    '— 11) Timestamp & keywords table —
    ActiveWorkbook.Worksheets("Cover").Range("cUpdateDate").Value = Date
    UpdateKeywordsTable

    '— 12) Refresh all —
    ActiveWorkbook.RefreshAll

    MsgBox "tbFiles synchronized.", vbInformation

CleanUp:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "FolderScan failed: " & Err.Description, vbCritical
End Sub



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

Private Function BuildJSONFromTableIgnoreYellow() As String
    Dim lo As ListObject, ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects("tbFiles")
        On Error GoTo 0
        If Not lo Is Nothing Then Exit For
    Next ws
    If lo Is Nothing Then Err.Raise vbObjectError + 1, , "tbFiles not found"

    If lo.DataBodyRange Is Nothing Then
        BuildJSONFromTableIgnoreYellow = "[]"
        Exit Function
    End If

    Dim json As String, sep As String
    Dim r As ListRow

    json = "[": sep = ""
    For Each r In lo.ListRows
        ' skip yellow rows (missing in filesystem)
        'If r.Range.Cells(1, lo.ListColumns("#").Index).Value = 78 Then
        '    Debug.Print "Break here"
        'End If
        If r.Range.Cells(1, lo.ListColumns("#").Index).Interior.Color = vbYellow Then
            ' skip
        Else
            With r.Range
                json = json & sep & "{"
                json = json & """RelativePath"":""" & j(.Cells(1, lo.ListColumns("RelativePath").Index).Value) & ""","
                json = json & """ObjectType"":""" & j(.Cells(1, lo.ListColumns("Object Type").Index).Value) & ""","
                json = json & """Description"":""" & j(.Cells(1, lo.ListColumns("Description").Index).Value) & ""","
                json = json & """Keywords"":""" & j(.Cells(1, lo.ListColumns("Keywords").Index).Value) & ""","
                json = json & """Reference"":""" & j(.Cells(1, lo.ListColumns("Reference").Index).Value) & ""","
                json = json & """DateDoc"":""" & j(.Cells(1, lo.ListColumns("Date Doc").Index).Value) & ""","
                json = json & """Approved"":""" & j(.Cells(1, lo.ListColumns("Approved?").Index).Value) & ""","
                json = json & """Link"":""" & j(.Cells(1, lo.ListColumns("Link").Index).Value) & """"
                json = json & "}"
            End With
            sep = ","
        End If
    Next r

    json = json & "]"
    BuildJSONFromTableIgnoreYellow = json
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
    Dim ws As Worksheet, loConfig As ListObject, loFiles As ListObject
    Dim templatePath As String, outputPath As String, repoName As String
    Dim html As String, json As String
    Dim stm As Object ' ADODB.Stream late-bound
    Dim found As Range

    ' 1) tbConfig
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Set loConfig = ws.ListObjects("tbConfig")
        On Error GoTo 0
        If Not loConfig Is Nothing Then Exit For
    Next ws
    If loConfig Is Nothing Then
        MsgBox "Configuration table 'tbConfig' not found.", vbCritical: Exit Sub
    End If

    ' Repository name
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
        MsgBox "No 'Html Template' in tbConfig.", vbCritical: Exit Sub
    End If
    templatePath = CStr(found.Offset(0, _
      loConfig.ListColumns("Value").Index - loConfig.ListColumns("Key").Index).Value)

    ' Html Index file
    With loConfig.DataBodyRange
        Set found = .Columns(loConfig.ListColumns("Key").Index) _
            .Find(What:="Html Index file", LookIn:=xlValues, LookAt:=xlWhole)
    End With
    If found Is Nothing Then
        MsgBox "No 'Html Index file' in tbConfig.", vbCritical: Exit Sub
    End If
    outputPath = CStr(found.Offset(0, _
      loConfig.ListColumns("Value").Index - loConfig.ListColumns("Key").Index).Value)

    ' 2) Read template
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2 ' text
        .Charset = "utf-8"
        .Open
        .LoadFromFile templatePath
        html = .ReadText(-1)
        .Close
    End With

    ' 3) Build JSON from tbFiles, ignoring yellow rows
    json = BuildJSONFromTableIgnoreYellow()

    ' 4) Replace REPNAME and timestamp
    html = Replace(html, "REPNAME", repoName, , , vbTextCompare)
    Dim stamp As String
    stamp = Format(Now, "d mmm HH:nn")
    html = Replace(html, "1 Jan 23:45", stamp, , , vbTextCompare)

    ' 5) Replace const fileData = [ ... ];
    Dim startPos As Long, openPos As Long, endPos As Long
    Dim before As String, after As String
    startPos = InStr(1, html, "const fileData =", vbTextCompare)
    If startPos > 0 Then
        openPos = InStr(startPos, html, "[")
        endPos = InStr(openPos, html, "];")
        If openPos > 0 And endPos > 0 Then
            before = Left$(html, startPos - 1)          ' <-- fix here
            after = Mid$(html, endPos + 2)              ' keep content after ];
            html = before & "const fileData = " & json & ";" & after
        End If
    End If
    
    ' 6) Write output utf-8
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


