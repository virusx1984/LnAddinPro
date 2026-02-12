' Define a simple Type to hold our data temporarily
Private Type LinkData
    Address As String
    category As String ' "External", "LNF", "Internal"
    Formula As String
    CellObj As Range
End Type



'+==========================================================+
'|                                                          |
'|                        <-- SECTION END -->               |
'|                                                          |
'+==========================================================+

' Purpose: Applies "Business Blue" formatting with conditional header lines.
' @param nHeaderRows: Number of rows to treat as header.
' @return: None
Public Sub ApplyTableStyle(nHeaderRows As Long)
    Dim rngFull As Range
    Dim rngHeader As Range
    Dim rngBody As Range
    Dim checkCell As Range
    Dim r As Long, c As Long
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    ' 1. Determine Target Range
    If Selection.Cells.CountLarge = 1 Then
        Set rngFull = Selection.CurrentRegion
    Else
        Set rngFull = Selection
    End If
    
    ' Validation
    If rngFull.Rows.count <= nHeaderRows Then
        MsgBox "Range is too small for " & nHeaderRows & " header rows.", vbExclamation
        GoTo ExitHandler
    End If
    
    ' 2. Define Header and Body Ranges
    Set rngHeader = rngFull.Resize(nHeaderRows, rngFull.Columns.count)
    Set rngBody = rngFull.offset(nHeaderRows, 0).Resize(rngFull.Rows.count - nHeaderRows, rngFull.Columns.count)
    
    ' ---------------------------------------------------------
    ' 3. APPLY HEADER STYLES
    ' ---------------------------------------------------------
    With rngHeader
        ' Basic Formatting
        .Interior.Color = RGB(0, 112, 192) ' Business Blue
        .Font.Color = vbWhite
        .Font.bold = True
        .Font.Name = "Arial"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        ' Clear all existing borders first
        .Borders.LineStyle = xlNone
        
        ' A. Vertical Separators (Always White)
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Color = vbWhite
            .Weight = xlThin
        End With
    End With
    
    ' B. Conditional Horizontal Separators
    ' Only draw white line if the cell BELOW is not empty.
    If nHeaderRows > 1 Then
        For r = 1 To nHeaderRows - 1
            For c = 1 To rngHeader.Columns.count
                If Len(Trim(rngHeader.Cells(r + 1, c).Value)) > 0 Then
                    With rngHeader.Cells(r, c).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Color = vbWhite
                        .Weight = xlThin
                    End With
                End If
            Next c
        Next r
    End If
    
    ' C. Header Bottom Border (Divider between Header and Body)
    With rngHeader.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(217, 217, 217) ' Light Gray
        .Weight = xlThin
    End With
    
    ' ---------------------------------------------------------
    ' 4. APPLY BODY STYLES
    ' ---------------------------------------------------------
    With rngBody
        .Interior.pattern = xlNone
        .Borders.LineStyle = xlNone
        
        ' A. Inner Horizontal Gray Lines (Between rows)
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Color = RGB(217, 217, 217) ' Light Gray
            .Weight = xlThin
        End With
        
        ' B. Bottom Horizontal Gray Line (For the last row) -> ADDED THIS
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(217, 217, 217) ' Light Gray
            .Weight = xlThin
        End With
    End With
    
    ' ---------------------------------------------------------
    ' 5. SMART ALIGNMENT
    ' ---------------------------------------------------------
    For c = 1 To rngBody.Columns.count
        Set checkCell = rngBody.Cells(1, c)
        
        ' Logic: Date=Left, Number=Right, Text=Left
        If IsDate(checkCell.Value) Then
            rngBody.Columns(c).HorizontalAlignment = xlLeft
        ElseIf IsNumeric(checkCell.Value) And Not IsEmpty(checkCell.Value) Then
            rngBody.Columns(c).HorizontalAlignment = xlRight
        Else
            rngBody.Columns(c).HorizontalAlignment = xlLeft
        End If
    Next c
    
    ' ---------------------------------------------------------
    ' 6. LAYOUT (FreezePanes Removed)
    ' ---------------------------------------------------------
    rngFull.Columns.AutoFit
    
    ' Cap width at 50
    For c = 1 To rngFull.Columns.count
        If rngFull.Columns(c).ColumnWidth > 50 Then
            rngFull.Columns(c).ColumnWidth = 50
            rngFull.Columns(c).WrapText = True
        Else
            rngFull.Columns(c).WrapText = False
        End If
    Next c
    
    ' Row Heights
    rngHeader.RowHeight = 20
    rngBody.RowHeight = 16.5
    
    ' Reset selection to top-left
    rngFull.Cells(1, 1).Select

ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Error in ApplyTableStyle: " & Err.Description, vbExclamation
    Resume ExitHandler
End Sub

'+==========================================================+
'|                                                          |
'|                        <-- SECTION END -->               |
'|                                                          |
'+==========================================================+
' Purpose: Convert range to Markdown Table format
' @param rng: Source Range
' @return: String containing Markdown text
Public Function RangeToMarkdown(rng As Range) As String
    Dim r As Long, c As Long
    Dim strBuilder As String
    Dim rowStr As String
    
    ' 1. Header
    rowStr = "|"
    For c = 1 To rng.Columns.count
        ' CHANGE: Use .Text instead of array value
        rowStr = rowStr & " " & CleanText(rng.Cells(1, c).Text) & " |"
    Next c
    strBuilder = strBuilder & rowStr & vbCrLf
    
    ' 2. Separator
    rowStr = "|"
    For c = 1 To rng.Columns.count
        rowStr = rowStr & " --- |"
    Next c
    strBuilder = strBuilder & rowStr & vbCrLf
    
    ' 3. Data
    For r = 2 To rng.Rows.count
        rowStr = "|"
        For c = 1 To rng.Columns.count
            ' CHANGE: Use .Text
            rowStr = rowStr & " " & CleanText(rng.Cells(r, c).Text) & " |"
        Next c
        strBuilder = strBuilder & rowStr & vbCrLf
    Next r
    
    RangeToMarkdown = strBuilder
End Function

' Purpose: Convert range to HTML Table format (Bootstrap ready)
' @param rng: Source Range
' @param includeClass: Boolean, if true adds Bootstrap classes
' @return: String containing HTML text
Public Function RangeToHTML(rng As Range, includeClass As Boolean) As String
    Dim r As Long, c As Long
    Dim strBuilder As String
    Dim tableClass As String
    
    If includeClass Then tableClass = " class=""table table-striped table-bordered"""
    strBuilder = "<table" & tableClass & ">" & vbCrLf
    
    ' 1. Header
    strBuilder = strBuilder & "  <thead>" & vbCrLf & "    <tr>" & vbCrLf
    For c = 1 To rng.Columns.count
        ' CHANGE: Use .Text
        strBuilder = strBuilder & "      <th>" & CleanText(rng.Cells(1, c).Text) & "</th>" & vbCrLf
    Next c
    strBuilder = strBuilder & "    </tr>" & vbCrLf & "  </thead>" & vbCrLf
    
    ' 2. Data
    strBuilder = strBuilder & "  <tbody>" & vbCrLf
    For r = 2 To rng.Rows.count
        strBuilder = strBuilder & "    <tr>" & vbCrLf
        For c = 1 To rng.Columns.count
            ' CHANGE: Use .Text
            strBuilder = strBuilder & "      <td>" & CleanText(rng.Cells(r, c).Text) & "</td>" & vbCrLf
        Next c
        strBuilder = strBuilder & "    </tr>" & vbCrLf
    Next r
    strBuilder = strBuilder & "  </tbody>" & vbCrLf & "</table>"
    
    RangeToHTML = strBuilder
End Function

' Helper: Clean text to prevent breaking table structure
Private Function CleanText(txt As String) As String
    Dim s As String
    s = txt
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, "|", "/")
    CleanText = Trim(s)
End Function
'+==========================================================+
'|                                                          |
'|                        <-- SECTION END -->               |
'|                                                          |
'+==========================================================+





' Purpose: Converts an Excel range into a compact, split-oriented JSON string format.
'          The output contains separate arrays for headers, types, and data rows to optimize performance.
'
' @param rng (Range): REQUIRED. The source data range to convert.
'                     - Row 1: Must contain unique Header names.
'                     - Row 2+: Contains the Data values.
' @param typeDict (Object): REQUIRED. A Scripting.Dictionary defining the column types.
'                           - Key (String): Header name matching the range.
'                           - Value (String): Type identifier ("cate"/"category" or "agg"/"aggregation").
'                           - Behavior: Defaults to "aggregation" if missing.
'                           - Constraint: At least one column MUST be defined as "category".
' @param formatDict (Object): OPTIONAL. A Scripting.Dictionary defining number formats for specific columns.
'                             - Key (String): Header name.
'                             - Value (String): VBA Format string (e.g., "0.00", "$#,##0", "yyyy-mm-dd").
'
' @return (String): A JSON formatted string containing "columns", "col_types", and "data" keys.
Public Function RangeToCompactJson( _
    rng As Range, _
    typeDict As Object, _
    Optional formatDict As Object _
) As String
    
    Dim dataArr As Variant
    Dim r As Long, c As Long
    Dim rowCount As Long, colCount As Long
    
    ' Buffers for string joining
    Dim headerList() As String
    Dim typeList() As String
    Dim rowList() As String
    Dim cellList() As String
    
    ' Performance Cache
    Dim colFormats() As String
    
    Dim headerName As String
    Dim dictValue As String
    Dim finalType As String
    Dim hasCategory As Boolean
    Dim cellValue As Variant
    Dim fmtStr As String
    Dim formattedVal As String
    
    ' Load range data into memory array
    dataArr = rng.Value2
    rowCount = UBound(dataArr, 1)
    colCount = UBound(dataArr, 2)
    
    ' Initialize arrays
    ReDim headerList(1 To colCount)
    ReDim typeList(1 To colCount)
    ReDim colFormats(1 To colCount) ' Cache formats per column
    
    If rowCount > 1 Then
        ReDim rowList(1 To rowCount - 1)
    Else
        RangeToCompactJson = "{}"
        Exit Function
    End If
    
    hasCategory = False
    
    ' --- Step 1: Process Headers, Types and Cache Formats ---
    For c = 1 To colCount
        headerName = CStr(dataArr(1, c))
        headerList(c) = """" & EscapeJson(headerName) & """"
        
        ' 1. Determine Type
        finalType = "aggregation"
        If Not typeDict Is Nothing Then
            If typeDict.Exists(headerName) Then
                dictValue = LCase(typeDict(headerName))
                If InStr(1, dictValue, "cate") > 0 Then
                    finalType = "category"
                    hasCategory = True
                ElseIf InStr(1, dictValue, "agg") > 0 Then
                    finalType = "aggregation"
                End If
            End If
        End If
        typeList(c) = """" & finalType & """"
        
        ' 2. Cache Format String (if exists)
        colFormats(c) = ""
        If Not formatDict Is Nothing Then
            If formatDict.Exists(headerName) Then
                colFormats(c) = formatDict(headerName)
            End If
        End If
    Next c
    
    ' --- Validation ---
    If Not hasCategory Then
        Err.Raise vbObjectError + 513, "RangeToCompactJson", _
            "Invalid Configuration: At least one column must be defined as 'category' (cate)."
    End If
    
    ' --- Step 2: Process Data Rows ---
    For r = 2 To rowCount
        ReDim cellList(1 To colCount)
        For c = 1 To colCount
            cellValue = dataArr(r, c)
            fmtStr = colFormats(c)
            
            ' Logic: Apply format if present, otherwise default handling
            If fmtStr <> "" And Not IsError(cellValue) And Not IsEmpty(cellValue) Then
                formattedVal = Format(cellValue, fmtStr)
                
                ' If formatted value looks like a pure number (no commas/symbols), treat as JSON Number
                ' Otherwise treat as JSON String (e.g., "$1,000" or "2026-01-01")
                If IsNumeric(formattedVal) And InStr(formattedVal, ",") = 0 And InStr(formattedVal, " ") = 0 Then
                    cellList(c) = formattedVal
                Else
                    cellList(c) = """" & EscapeJson(formattedVal) & """"
                End If
            Else
                ' Default Handling (No specific format requested)
                If IsNumeric(cellValue) And Not IsEmpty(cellValue) Then
                    cellList(c) = CStr(cellValue)
                Else
                    If IsError(cellValue) Then cellValue = "Error"
                    If IsEmpty(cellValue) Then cellValue = ""
                    cellList(c) = """" & EscapeJson(CStr(cellValue)) & """"
                End If
            End If
        Next c
        rowList(r - 1) = "    [" & Join(cellList, ", ") & "]"
    Next r
    
    ' --- Step 3: Assemble Final JSON ---
    Dim jsonParts(1 To 3) As String
    jsonParts(1) = "  ""columns"": [" & Join(headerList, ", ") & "]"
    jsonParts(2) = "  ""col_types"": [" & Join(typeList, ", ") & "]"
    jsonParts(3) = "  ""data"": [" & vbCrLf & Join(rowList, "," & vbCrLf) & vbCrLf & "  ]"
    
    RangeToCompactJson = "{" & vbCrLf & Join(jsonParts, "," & vbCrLf) & vbCrLf & "}"

End Function

Private Function EscapeJson(s As String) As String
    Dim temp As String
    temp = Replace(s, "\", "\\")
    temp = Replace(temp, """", "\""")
    temp = Replace(temp, vbCrLf, "\n")
    temp = Replace(temp, vbCr, "\r")
    temp = Replace(temp, vbLf, "\n")
    temp = Replace(temp, vbTab, "\t")
    EscapeJson = temp
End Function
'+==========================================================+
'|                                                          |
'|                        <-- SECTION END -->               |
'|                                                          |
'+==========================================================+






' Purpose: Compares two Excel data ranges based on specified column categories.
' UPDATED: Accepts custom Table Names for the output header.
Public Function CompareExcelRanges( _
    ByVal rng1 As Range, _
    ByVal rng2 As Range, _
    ByVal indexCols As Variant, _
    Optional ByVal ignoreCols As Variant, _
    Optional ByVal referenceCols As Variant, _
    Optional ByVal referenceColDirections As Variant, _
    Optional ByVal explicitCompareCols As Variant, _
    Optional ByVal tableName1 As String = "Table1", _
    Optional ByVal tableName2 As String = "Table2", _
    Optional ByVal flattenHeader As Boolean = False _
) As Variant

    ' --- Variable Declarations ---
    Dim colCount As Long
    Dim i As Long, j As Long
    Dim header1() As Variant
    
    ' Dictionaries
    Dim dictHeaders As Object
    Dim dictT1 As Object
    Dim dictT2 As Object
    Dim indexColIndexes As Object
    Dim compareColIndexes As Object
    Dim refColIndexes As Object
    
    ' Internal String Arrays
    Dim arrIndex() As String
    Dim arrIgnore() As String
    Dim arrRef() As String
    Dim arrComp() As String
    
    ' Collection
    Dim resultCollection As Collection
    
    ' Control variables
    Dim key As Variant
    Dim compHeader As Variant
    Dim colName As Variant
    Dim arrTemp As Variant
    
    ' Direction Logic Variable
    Dim actualRefDirs As Object
    
    ' Initialization
    Set dictHeaders = CreateObject("Scripting.Dictionary")
    Set dictT1 = CreateObject("Scripting.Dictionary")
    Set dictT2 = CreateObject("Scripting.Dictionary")
    Set indexColIndexes = CreateObject("Scripting.Dictionary")
    Set compareColIndexes = CreateObject("Scripting.Dictionary")
    Set refColIndexes = CreateObject("Scripting.Dictionary")
    Set resultCollection = New Collection
    
    On Error GoTo ErrorHandler
    
    ' --- 1. Input Validation and Handling Optionals ---
    
    ' 1a. Validate indexCols (Required)
    If Not VBA.IsArray(indexCols) Then
        If IsEmpty(indexCols) Or CStr(indexCols) = "" Then GoTo IndexErr
        CompareExcelRanges = Array(Array("Error: indexCols must be passed as an Array."))
        Exit Function
    End If
    arrTemp = indexCols
    ReDim arrIndex(LBound(arrTemp) To UBound(arrTemp))
    For i = LBound(arrTemp) To UBound(arrTemp): arrIndex(i) = Trim(CStr(arrTemp(i))): Next i
    GoTo CheckIgnore

IndexErr:
    CompareExcelRanges = Array(Array("Error: indexCols is REQUIRED."))
    Exit Function

CheckIgnore:
    ' 1b. Handle ignoreCols
    If IsMissing(ignoreCols) Or Not VBA.IsArray(ignoreCols) Then
        arrIgnore = Split(vbNullString)
    Else
        arrTemp = ignoreCols
        ReDim arrIgnore(LBound(arrTemp) To UBound(arrTemp))
        For i = LBound(arrTemp) To UBound(arrTemp): arrIgnore(i) = Trim(CStr(arrTemp(i))): Next i
    End If

    ' 1c. Handle referenceCols
    If IsMissing(referenceCols) Or Not VBA.IsArray(referenceCols) Then
        arrRef = Split(vbNullString)
    Else
        arrTemp = referenceCols
        ReDim arrRef(LBound(arrTemp) To UBound(arrTemp))
        For i = LBound(arrTemp) To UBound(arrTemp): arrRef(i) = Trim(CStr(arrTemp(i))): Next i
    End If

    ' 1d. Handle referenceColDirections
    Set actualRefDirs = Nothing
    If Not IsMissing(referenceColDirections) Then
        If IsObject(referenceColDirections) Then Set actualRefDirs = referenceColDirections
    End If
    
    ' 1e. Handle explicitCompareCols
    Dim useExplicitCompare As Boolean: useExplicitCompare = False
    If Not IsMissing(explicitCompareCols) Then
        If VBA.IsArray(explicitCompareCols) Then
            arrTemp = explicitCompareCols
            If UBound(arrTemp) >= LBound(arrTemp) Then
                ReDim arrComp(LBound(arrTemp) To UBound(arrTemp))
                For i = LBound(arrTemp) To UBound(arrTemp): arrComp(i) = Trim(CStr(arrTemp(i))): Next i
                useExplicitCompare = True
            End If
        End If
    End If

    ' --- 2. Initial Checks (Column Count and Header Matching) ---
    colCount = rng1.Columns.count
    If colCount <> rng2.Columns.count Then
        CompareExcelRanges = Array(Array("Error: The two ranges have different column counts."))
        Exit Function
    End If
    
    header1 = rng1.Rows(1).Value
    Dim header2() As Variant
    header2 = rng2.Rows(1).Value
    
    For i = 1 To colCount
        If StrComp(CStr(header1(1, i)), CStr(header2(1, i)), vbTextCompare) <> 0 Then
            CompareExcelRanges = Array(Array("Error: Column headers do not match in name or order."))
            Exit Function
        End If
        dictHeaders.Add LCase(Trim(CStr(header1(1, i)))), i
    Next i
    
    ' --- 3. Parameter Validation & Column Categorization ---
    
    ' 3a. Index Columns
    For i = LBound(arrIndex) To UBound(arrIndex)
        colName = arrIndex(i)
        If dictHeaders.Exists(LCase(colName)) Then
            If Not indexColIndexes.Exists(colName) Then
                indexColIndexes.Add colName, dictHeaders(LCase(colName))
            End If
        Else
            CompareExcelRanges = Array(Array("Error: Index column '" & colName & "' not found."))
            Exit Function
        End If
    Next i
    
    ' 3b. Reference Columns
    If (UBound(arrRef) - LBound(arrRef) + 1) > 0 Then
        For i = LBound(arrRef) To UBound(arrRef)
            colName = arrRef(i)
            If dictHeaders.Exists(LCase(colName)) Then
                If Not refColIndexes.Exists(colName) Then
                    refColIndexes.Add colName, dictHeaders(LCase(colName))
                End If
            End If
        Next i
    End If
    
    ' Overlap check
    For Each compHeader In indexColIndexes.Keys
        If refColIndexes.Exists(compHeader) Then
            CompareExcelRanges = Array(Array("Error: Column '" & compHeader & "' is in both Index and Ref."))
            Exit Function
        End If
    Next
    
    ' 3c. Compare Columns
    If useExplicitCompare Then
        For i = LBound(arrComp) To UBound(arrComp)
            colName = arrComp(i)
            If dictHeaders.Exists(LCase(colName)) Then
                If Not compareColIndexes.Exists(colName) Then
                    compareColIndexes.Add colName, dictHeaders(LCase(colName))
                End If
            End If
        Next i
    Else
        For i = 1 To colCount
            Dim currentHeader As String: currentHeader = Trim(CStr(header1(1, i)))
            Dim lowerHeader As String: lowerHeader = LCase(currentHeader)
            
            Dim isIndex As Boolean: isIndex = indexColIndexes.Exists(currentHeader)
            Dim isRef As Boolean: isRef = refColIndexes.Exists(currentHeader)
            Dim isIgnore As Boolean: isIgnore = False
            
            If (UBound(arrIgnore) - LBound(arrIgnore) + 1) > 0 Then
                For Each colName In arrIgnore
                    If StrComp(lowerHeader, LCase(Trim(CStr(colName))), vbTextCompare) = 0 Then
                        isIgnore = True
                        Exit For
                    End If
                Next
            End If
            
            If Not isIndex And Not isRef And Not isIgnore Then
                compareColIndexes.Add currentHeader, i
            End If
        Next i
    End If
    
    ' --- 4. Build Dictionaries from Ranges ---
    Call PopulateDictionary(rng1, indexColIndexes, dictT1)
    Call PopulateDictionary(rng2, indexColIndexes, dictT2)

    ' --- 5. Comparison Logic ---
    Dim allKeys As Object: Set allKeys = CreateObject("Scripting.Dictionary")
    
    For Each key In dictT1.Keys
        If Not allKeys.Exists(key) Then allKeys.Add key, True
    Next key
    For Each key In dictT2.Keys
        If Not allKeys.Exists(key) Then allKeys.Add key, True
    Next key
    
    For Each key In allKeys.Keys
        Dim valuesT1 As Variant: valuesT1 = dictT1.item(key)
        Dim valuesT2 As Variant: valuesT2 = dictT2.item(key)
        
        Dim isT1Present As Boolean: isT1Present = Not IsEmpty(valuesT1)
        Dim isT2Present As Boolean: isT2Present = Not IsEmpty(valuesT2)
        
        Dim status As String
        Dim hasDiff As Boolean: hasDiff = False
        
        If isT1Present And isT2Present Then
            For Each compHeader In compareColIndexes.Keys
                i = compareColIndexes.item(compHeader)
                If CStr(valuesT1(i)) <> CStr(valuesT2(i)) Then
                    hasDiff = True
                    Exit For
                End If
            Next
        Else
            hasDiff = True
        End If

        ' Set Status
        If isT1Present And isT2Present Then
            status = "Both"
        ElseIf isT1Present Then
            status = "Left Only"
        ElseIf isT2Present Then
            status = "Right Only"
        End If
        
'        If Not hasDiff And (isT1Present And isT2Present) Then GoTo NextKey

        Dim rowResult As Object: Set rowResult = CreateObject("Scripting.Dictionary")
        rowResult.Add "Status", status
        
        ' 2b. Index Columns
        For Each compHeader In indexColIndexes.Keys
            i = indexColIndexes.item(compHeader)
            If isT1Present Then
                rowResult.Add compHeader, valuesT1(i)
            Else
                rowResult.Add compHeader, valuesT2(i)
            End If
        Next
        
        ' 2c. Reference Columns
        For Each compHeader In refColIndexes.Keys
            i = refColIndexes.item(compHeader)
            Dim refVal As Variant: refVal = ""
            Dim refT2 As Variant: If isT2Present Then refT2 = valuesT2(i) Else refT2 = Empty
            Dim refT1 As Variant: If isT1Present Then refT1 = valuesT1(i) Else refT1 = Empty
            
            Dim prioritizeT2 As Boolean: prioritizeT2 = True
            
            If Not actualRefDirs Is Nothing Then
                If actualRefDirs.Exists(compHeader) Then
                    prioritizeT2 = CBool(actualRefDirs(compHeader))
                End If
            End If
            
            If prioritizeT2 Then
                If Not IsEmpty(refT2) And refT2 <> "" Then refVal = refT2
                If (IsEmpty(refVal) Or refVal = "") And (Not IsEmpty(refT1) And refT1 <> "") Then refVal = refT1
            Else
                If Not IsEmpty(refT1) And refT1 <> "" Then refVal = refT1
                If (IsEmpty(refVal) Or refVal = "") And (Not IsEmpty(refT2) And refT2 <> "") Then refVal = refT2
            End If
            
            rowResult.Add compHeader & "_Ref", refVal
        Next
        
        ' 2d. Comparison Columns
        Dim t1Values As Object: Set t1Values = CreateObject("Scripting.Dictionary")
        Dim t2Values As Object: Set t2Values = CreateObject("Scripting.Dictionary")
        Dim diffValues As Object: Set diffValues = CreateObject("Scripting.Dictionary")
        
        For Each compHeader In compareColIndexes.Keys
            i = compareColIndexes.item(compHeader)
            
            Dim val1 As Variant
            If isT1Present Then val1 = valuesT1(i) Else val1 = 0
            
            Dim val2 As Variant
            If isT2Present Then val2 = valuesT2(i) Else val2 = 0
            
            Dim diffVal As Variant: diffVal = ""
            If IsNumeric(val1) And IsNumeric(val2) Then
                diffVal = CDbl(val2) - CDbl(val1)
            End If
            
            t1Values.Add compHeader, val1
            t2Values.Add compHeader, val2
            diffValues.Add compHeader, diffVal
        Next
        
        For Each compHeader In t1Values.Keys: rowResult.Add compHeader & "_T1", t1Values.item(compHeader): Next
        For Each compHeader In t2Values.Keys: rowResult.Add compHeader & "_T2", t2Values.item(compHeader): Next
        For Each compHeader In diffValues.Keys: rowResult.Add compHeader & "_Diff", diffValues.item(compHeader): Next
        
        resultCollection.Add rowResult.Items
NextKey:
    Next key
    
    ' --- 6. Assemble Output (FIXED) ---
    If resultCollection.count = 0 Then
        CompareExcelRanges = Array(Array("Success: No differences found."))
        Exit Function
    End If
    
    Dim totalOutputCols As Long
    totalOutputCols = 1 + indexColIndexes.count + refColIndexes.count + (compareColIndexes.count * 3)
    
    Dim arrResult() As Variant
    Dim headerRows As Long
    
    ' Determine number of header rows based on user preference
    If flattenHeader Then
        headerRows = 1
    Else
        headerRows = 2
    End If
    
    ' Resize result array to fit headers + data
    ReDim arrResult(1 To resultCollection.count + headerRows, 1 To totalOutputCols)
    
    Dim outputRow As Long
    Dim colIndex As Long
    ' [DELETED] Dim colName As Variant   <-- Caused Error: Already declared at top
    ' [DELETED] Dim compHeader As Variant <-- Caused Error: Already declared at top
    
    If flattenHeader Then
        ' === FLAT HEADER LOGIC (1 ROW) ===
        ' Merges TableName and ColumnName using underscore (e.g., T1_Price)
        outputRow = 1
        colIndex = 1
        
        ' 1. Status Column
        arrResult(outputRow, colIndex) = "Status": colIndex = colIndex + 1
        
        ' 2. Index Columns
        For Each colName In indexColIndexes.Keys
            arrResult(outputRow, colIndex) = colName
            colIndex = colIndex + 1
        Next
        
        ' 3. Reference Columns
        For Each colName In refColIndexes.Keys
            arrResult(outputRow, colIndex) = colName & "_Ref"
            colIndex = colIndex + 1
        Next
        
        ' 4. Compare Columns (Flattened)
        ' Group 1: Table 1 Values
        For Each compHeader In compareColIndexes.Keys
            arrResult(outputRow, colIndex) = tableName1 & "_" & compHeader
            colIndex = colIndex + 1
        Next
        ' Group 2: Table 2 Values
        For Each compHeader In compareColIndexes.Keys
            arrResult(outputRow, colIndex) = tableName2 & "_" & compHeader
            colIndex = colIndex + 1
        Next
        ' Group 3: Difference Values
        For Each compHeader In compareColIndexes.Keys
            arrResult(outputRow, colIndex) = "Diff_" & compHeader
            colIndex = colIndex + 1
        Next
        
    Else
        ' === STANDARD LOGIC (2 ROWS) ===
        ' Row 1: Group Names (Table1, Table2, Diff)
        ' Row 2: Column Names
        
        ' Header Row A (Top Level)
        outputRow = 1
        colIndex = 1
        
        arrResult(outputRow, colIndex) = "": colIndex = colIndex + 1
        For i = 1 To indexColIndexes.count: arrResult(outputRow, colIndex) = "": colIndex = colIndex + 1: Next
        For i = 1 To refColIndexes.count: arrResult(outputRow, colIndex) = "Ref": colIndex = colIndex + 1: Next
        For i = 1 To compareColIndexes.count: arrResult(outputRow, colIndex) = tableName1: colIndex = colIndex + 1: Next
        For i = 1 To compareColIndexes.count: arrResult(outputRow, colIndex) = tableName2: colIndex = colIndex + 1: Next
        For i = 1 To compareColIndexes.count: arrResult(outputRow, colIndex) = "Diff": colIndex = colIndex + 1: Next
        
        ' Header Row B (Column Names)
        outputRow = 2
        colIndex = 1
        arrResult(outputRow, colIndex) = "Status": colIndex = colIndex + 1
        For Each colName In indexColIndexes.Keys: arrResult(outputRow, colIndex) = colName: colIndex = colIndex + 1: Next
        For Each colName In refColIndexes.Keys: arrResult(outputRow, colIndex) = colName: colIndex = colIndex + 1: Next
        For Each compHeader In compareColIndexes.Keys: arrResult(outputRow, colIndex) = compHeader: colIndex = colIndex + 1: Next
        For Each compHeader In compareColIndexes.Keys: arrResult(outputRow, colIndex) = compHeader: colIndex = colIndex + 1: Next
        For Each compHeader In compareColIndexes.Keys: arrResult(outputRow, colIndex) = compHeader: colIndex = colIndex + 1: Next
    End If
    
    ' === Data Rows ===
    ' Start outputting data after the header rows
    outputRow = headerRows + 1
    
    Dim dataArray As Variant
    For Each dataArray In resultCollection
        colIndex = 1
        For i = LBound(dataArray) To UBound(dataArray)
            arrResult(outputRow, colIndex) = dataArray(i)
            colIndex = colIndex + 1
        Next
        outputRow = outputRow + 1
    Next
    
    Set resultCollection = Nothing
    CompareExcelRanges = arrResult
    Exit Function

ErrorHandler:
    CompareExcelRanges = Array(Array("Error: Unexpected runtime error. Number: " & Err.Number & ", Description: " & Err.Description))
End Function

' --- Helper for Dictionary Population (Must accompany main function) ---
Private Sub PopulateDictionary(ByVal rng As Range, ByVal indexCols As Object, ByRef targetDict As Object)
    Dim dataRange As Range
    Dim dataArray As Variant
    Dim r As Long, c As Long
    Dim colIndex As Variant
    Dim keyString As String
    Dim existingValues As Variant
    Dim newValues() As Variant
    Dim arrayColCount As Long
    
    Set dataRange = rng.offset(1, 0).Resize(rng.Rows.count - 1, rng.Columns.count)
    If dataRange.Rows.count = 0 Then Exit Sub
    
    dataArray = dataRange.Value
    arrayColCount = UBound(dataArray, 2)
    
    For r = 1 To UBound(dataArray, 1)
        keyString = ""
        For Each colIndex In indexCols.Items
            keyString = keyString & CStr(dataArray(r, colIndex)) & "|"
        Next colIndex
        If Len(keyString) > 0 Then keyString = Left(keyString, Len(keyString) - 1)
        
        If Not targetDict.Exists(keyString) Then
            ReDim newValues(1 To arrayColCount)
            For c = 1 To arrayColCount
                newValues(c) = dataArray(r, c)
            Next c
            targetDict.Add keyString, newValues
        Else
            existingValues = targetDict.item(keyString)
            For c = 1 To arrayColCount
                Dim isIndexCol As Boolean: isIndexCol = False
                For Each colIndex In indexCols.Items
                    If CLng(colIndex) = c Then
                        isIndexCol = True
                        Exit For
                    End If
                Next colIndex
                
                If Not isIndexCol Then
                    If IsNumeric(existingValues(c)) And IsNumeric(dataArray(r, c)) Then
                        existingValues(c) = CDbl(existingValues(c)) + CDbl(dataArray(r, c))
                    End If
                End If
            Next c
            targetDict.item(keyString) = existingValues
        End If
    Next r
End Sub
'+==========================================================+
'|                                                          |
'|                        <-- SECTION END -->               |
'|                                                          |
'+==========================================================+






' Purpose: Generates a sequential time series array (e.g., dates, quarter strings, or annual labels)
'          based on a composite string that defines the required frequency for a sequence of years.
'
' @param IntervalTypeCombo (String): A string where each character specifies the interval type for a year,
'                                   starting from StartYear. (e.g., "MMQ" means Year 1 is Monthly, Year 2 is Monthly, Year 3 is Quarterly).
'                                   M = Monthly, Q = Quarterly, H = Semi-Annually, Y = Annually.
' @param IncludeAnnualTotal (Boolean): If True, an annual label (e.g., "Y2025") is appended after the
'                                      periodic data (M, Q, or H) for that year.
' @param StartYear (Long): The starting year of the time series.
' @return (Variant): A one-dimensional array containing the generated time series elements (Dates or Strings).
Public Function GenerateTimeSeries( _
    ByVal IntervalTypeCombo As String, _
    ByVal IncludeAnnualTotal As Boolean, _
    ByVal StartYear As Long _
) As Variant

    ' Declare and initialize a Collection object to dynamically store the series elements.
    Dim resultList As New Collection
    Dim currentYear As Long
    Dim yearIntervalType As String
    Dim i As Long, j As Long, k As Long
    Dim seriesYear As Date ' Variable to hold the start date of the current year.

    ' Loop through the years determined by the length of the IntervalTypeCombo string.
    ' i represents the position in the combo string (and the year index).
    For i = 1 To Len(IntervalTypeCombo)
        currentYear = StartYear + i - 1 ' Calculate the current calendar year.
        yearIntervalType = Mid(IntervalTypeCombo, i, 1) ' Extract the frequency code for the current year.
        seriesYear = DateSerial(currentYear, 1, 1) ' Start date of the current year (unused for string types, but useful for date types).

        ' Determine the required time interval based on the character code.
        Select Case UCase(yearIntervalType)
            Case "M" ' Monthly Frequency
                ' Add monthly dates (first day of each month)
                For j = 1 To 12
                    resultList.Add DateSerial(currentYear, j, 1)
                Next j
                ' Add the annual total label if required.
                If IncludeAnnualTotal Then
                    resultList.Add "Y" & CStr(currentYear)
                End If

            Case "Q" ' Quarterly Frequency
                ' Add quarterly strings (e.g., "2025 Q1")
                For j = 1 To 4
                    ' The quarter string is added as a text label.
                    resultList.Add CStr(currentYear) & " Q" & CStr(j)
                Next j
                ' Add the annual total label if required.
                If IncludeAnnualTotal Then
                    resultList.Add "Y" & CStr(currentYear)
                End If

            Case "H" ' Semi-Annual Frequency
                ' Add semi-annual strings (e.g., "2025 H1")
                For j = 1 To 2
                    ' The half-year string is added as a text label.
                    resultList.Add CStr(currentYear) & " H" & CStr(j)
                Next j
                ' Add the annual total label if required.
                If IncludeAnnualTotal Then
                    resultList.Add "Y" & CStr(currentYear)
                End If

            Case "Y" ' Annual Frequency (Yearly only)
                ' Add only the annual label (e.g., "Y2025").
                ' No periodic data is added since the frequency is annual.
                resultList.Add "Y" & CStr(currentYear)

            Case Else
                ' If the character is not one of the defined interval types, it is simply skipped.
        End Select
    Next i

    ' --- Final Output Conversion ---
    
    ' Convert the dynamic Collection object into a static 1D Variant Array for output.
    Dim resultArray() As Variant
    Dim listCount As Long
    listCount = resultList.count

    If listCount > 0 Then
        ' ReDim the array to the exact size of the Collection (1-based index).
        ReDim resultArray(1 To listCount)
        For k = 1 To listCount
            ' Transfer elements from the Collection to the array.
            resultArray(k) = resultList(k)
        Next k
        GenerateTimeSeries = resultArray ' Return the filled array.
    Else
        GenerateTimeSeries = Array() ' Return an empty array if the Collection is empty.
    End If

End Function
'+==========================================================+
'|                                                          |
'|                        <-- SECTION END -->               |
'|                                                          |
'+==========================================================+


' Purpose: Melts a table (wide-to-long format transformation), similar to pandas.melt().
' It preserves index columns and transforms other columns into variable-value pairs.
' UPDATED: Automatically aligns idColumnsRange rows to match tableRange rows.
'
' @param tableRange (Range): The full table range, including headers, index, and value columns.
' @param idColumnsRange (Range): The range of columns to be used as index (id_vars).
' @param variableName (String): The header name for the "Variable" column in the output.
' @return (Variant): A 2D array with columns [index columns..., variableName, Value]. Returns #VALUE! error on failure.
Public Function MeltData(ByVal tableRange As Range, ByVal idColumnsRange As Range, ByVal variableName As String) As Variant
    
    ' --- Variable Declarations ---
    Dim data As Variant
    Dim result As Variant
    Dim headers As Variant
    
    Dim dataRow As Long, dataCol As Long
    Dim resultCol As Long
    Dim currentResultRow As Long
    
    Dim totalRows As Long, totalCols As Long
    Dim idColCount As Long
    Dim valueColCount As Long
    Dim resultRowCount As Long
    
    ' --- Data Input and Initial Checks ---
    ' Read the entire table data into a 2D array (fastest method).
    data = tableRange.Value
    
    ' Get dimensions, assuming data is 1-based (standard for Range.Value).
    totalRows = UBound(data, 1) ' Total number of rows including header
    totalCols = UBound(data, 2) ' Total number of columns
    
    ' ==============================================================================
    ' [UPDATED] Auto-Align ID Columns to Match Table Rows
    ' Logic: If rows don't match (e.g., A:A vs B2:E10), crop ID range to match Table.
    ' ==============================================================================
    If (tableRange.Rows.count <> idColumnsRange.Rows.count) Or _
       (tableRange.Row <> idColumnsRange.Row) Then
        
        ' Ensure both ranges are on the same worksheet
        If tableRange.Parent.Name = idColumnsRange.Parent.Name Then
            On Error Resume Next
            ' Intersect: Restrict ID columns to the rows covered by tableRange
            Set idColumnsRange = Application.Intersect(idColumnsRange.EntireColumn, tableRange.EntireRow)
            On Error GoTo 0
            
            ' Safety Check: If intersection failed (no overlap), return Error
            If idColumnsRange Is Nothing Then
                MeltData = CVErr(xlErrValue)
                Exit Function
            End If
        Else
            ' Cannot process ranges across different sheets
            MeltData = CVErr(xlErrValue)
            Exit Function
        End If
    End If
    ' ==============================================================================

    ' Get ID column count from the passed (and potentially adjusted) range.
    idColCount = idColumnsRange.Columns.count
    valueColCount = totalCols - idColCount
    
    ' Validate that there are value columns to melt.
    If valueColCount <= 0 Then
        ' The ID columns must be strictly less than the total columns.
        MeltData = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' --- Calculate Result Dimensions ---
    ' Result rows = (Data rows - 1 for header) * (Number of value columns) + 1 for header row
    resultRowCount = (totalRows - 1) * valueColCount + 1
    ' Result columns = ID columns + Variable column + Value column
    Dim resultColCount As Long
    resultColCount = idColCount + 2
    
    ' Initialize the result array
    ReDim result(1 To resultRowCount, 1 To resultColCount)
    
    ' Extract headers from the first row of the data array.
    headers = tableRange.Rows(1).Value
    
    ' --- Set Headers in the Result Array ---
    
    ' 1. Set ID Column Headers
    Dim idHeaders As Variant
    idHeaders = idColumnsRange.Rows(1).Value ' Read headers from the first row of ID columns

    ' --- FIX: Single column Range.Value trap ---
    ' If idColumnsRange has only one column, idHeaders is a single value, not an array.
    If idColCount = 1 Then
        Dim tempHeader(1 To 1, 1 To 1) As Variant
        tempHeader(1, 1) = idHeaders
        idHeaders = tempHeader ' Replace the single value with the 2D array
    End If
    ' -------------------------------------------

    For resultCol = 1 To idColCount
        ' Accessing idHeaders(1, resultCol) is now safe.
        result(1, resultCol) = idHeaders(1, resultCol)
    Next resultCol
    
    ' 2. Set Variable and Value Headers
    result(1, idColCount + 1) = variableName
    result(1, idColCount + 2) = "Value"
    
    ' --- Melt Data Transformation ---
    currentResultRow = 2 ' Start filling data from the second row

    ' Iterate through each data row (i.e., data(dataRow)) starting after the header (row 2).
    For dataRow = 2 To totalRows
        
        ' Iterate through each value column that needs to be melted.
        For dataCol = idColCount + 1 To totalCols
            
            ' 1. Copy ID Columns (data(dataRow, 1) to data(dataRow, idColCount))
            For resultCol = 1 To idColCount
                result(currentResultRow, resultCol) = data(dataRow, resultCol)
            Next resultCol
            
            ' 2. Assign Variable Name (Header of the current value column)
            result(currentResultRow, idColCount + 1) = headers(1, dataCol)
            
            ' 3. Assign Value
            result(currentResultRow, idColCount + 2) = data(dataRow, dataCol)
            
            ' Move to next result row
            currentResultRow = currentResultRow + 1
        Next dataCol
    Next dataRow
    
    ' --- Return the Array ---
    MeltData = result
    
End Function
'+==========================================================+
'|                                                          |
'|                        <-- SECTION END -->               |
'|                                                          |
'+==========================================================+