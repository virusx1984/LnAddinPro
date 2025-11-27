Attribute VB_Name = "mod_funcs"



' Purpose: Compares two Excel data ranges (Table 1 and Table 2) based on specified column categories (Index, Ignore, Reference).
'          It identifies mismatches and unique rows, and compiles the results into a structured 2D array.
'
' @param rng1 (Range): REQUIRED. The first data range (Table 1), which MUST include the header row.
' @param rng2 (Range): REQUIRED. The second data range (Table 2), which MUST include the header row. The headers must match rng1 exactly in name and order.
' @param indexCols (Variant): REQUIRED. A 1D Array of column names (Strings) used to create the unique lookup key for each row.
'                              These columns are always included in the output.
'                              Usage Example: Array("col1", "col2")
' @param ignoreCols (Variant): OPTIONAL. A 1D Array of column names (Strings) whose values will NOT be checked for differences.
'                              These columns are excluded from the final output array structure.
'                              Usage Example: Array("col4") or an empty array Array()
' @param referenceCols (Variant): OPTIONAL. A 1D Array of column names (Strings) that are neither Index nor Comparison fields.
'                                 Their values are filled using a VLOOKUP-like priority based on vlookupOrder.
'                                 Usage Example: Array("col3") or an empty array Array()
' @param vlookupOrder (Boolean): REQUIRED. Determines the fill priority for the referenceCols values when a match is found in both tables:
'                                - True: Priority is Table 2, then Table 1 (attempts to fill from T2 first).
'                                - False: Priority is Table 1, then Table 2 (attempts to fill from T1 first).
' @return (Variant): A two-dimensional, 1-based array containing the comparison results, including a two-row header structure.
Public Function CompareExcelRanges( _
    ByVal rng1 As Range, _
    ByVal rng2 As Range, _
    ByVal indexCols As Variant, _
    ByVal ignoreCols As Variant, _
    ByVal referenceCols As Variant, _
    ByVal vlookupOrder As Boolean _
) As Variant

    ' Core function to compare two Excel ranges and return the result as a 2D Variant array.

    ' --- Variable Declarations (Mixed Binding) ---
    Dim colCount As Long
    Dim i As Long, j As Long
    Dim header1() As Variant
    
    ' Dictionaries (Late Binding - Object Type)
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
    
    ' Collection (Early Binding - VBA native Collection)
    Dim resultCollection As Collection
    
    ' Control variables
    Dim key As Variant
    Dim compHeader As Variant
    Dim colName As Variant
    
    ' Initialization (Dictionary using CreateObject for Late Binding)
    Set dictHeaders = CreateObject("Scripting.Dictionary")
    Set dictT1 = CreateObject("Scripting.Dictionary")
    Set dictT2 = CreateObject("Scripting.Dictionary")
    Set indexColIndexes = CreateObject("Scripting.Dictionary")
    Set compareColIndexes = CreateObject("Scripting.Dictionary")
    Set refColIndexes = CreateObject("Scripting.Dictionary")
    Set resultCollection = New Collection
    
    On Error GoTo ErrorHandler
    
    ' --- 1. Input Validation and Array Conversion (Skipped for brevity, assume valid) ---
    Dim arrTemp As Variant
    
    If Not VBA.IsArray(indexCols) Then
        If IsEmpty(indexCols) Or indexCols = "" Then GoTo EmptyIndexCheck
        CompareExcelRanges = Array(Array("Error: indexCols must be passed as an Array (e.g., Array(""col1"", ""col2""))."))
        Exit Function
    End If
    arrTemp = indexCols
    ReDim arrIndex(LBound(arrTemp) To UBound(arrTemp))
    For i = LBound(arrTemp) To UBound(arrTemp): arrIndex(i) = Trim(CStr(arrTemp(i))): Next i
EmptyIndexCheck:
    
    If Not VBA.IsArray(ignoreCols) Then
        If IsEmpty(ignoreCols) Or ignoreCols = "" Then GoTo EmptyIgnoreCheck
        CompareExcelRanges = Array(Array("Error: ignoreCols must be passed as an Array (e.g., Array(""col4""))."))
        Exit Function
    End If
    arrTemp = ignoreCols
    ReDim arrIgnore(LBound(arrTemp) To UBound(arrTemp))
    For i = LBound(arrTemp) To UBound(arrTemp): arrIgnore(i) = Trim(CStr(arrTemp(i))): Next i
EmptyIgnoreCheck:

    If Not VBA.IsArray(referenceCols) Then
        If IsEmpty(referenceCols) Or referenceCols = "" Then GoTo EmptyRefCheck
        CompareExcelRanges = Array(Array("Error: referenceCols must be passed as an Array (e.g., Array(""col3""))."))
        Exit Function
    End If
    arrTemp = referenceCols
    ReDim arrRef(LBound(arrTemp) To UBound(arrTemp))
    For i = LBound(arrTemp) To UBound(arrTemp): arrRef(i) = Trim(CStr(arrTemp(i))): Next i
EmptyRefCheck:

    ' --- 2. Initial Checks (Column Count and Header Matching) ---
    colCount = rng1.Columns.Count
    If colCount <> rng2.Columns.Count Then
        CompareExcelRanges = Array(Array("Error: The two ranges have different column counts and cannot be compared."))
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
    
    ' --- 3. Parameter Validation & Column Categorization (Skipped for brevity, assume valid) ---
    Dim checkArrays As Variant
    checkArrays = Array(Array("Index", arrIndex, indexColIndexes), Array("Reference", arrRef, refColIndexes))
    
    For Each checkArr In checkArrays
        Dim arrType As String: arrType = checkArr(0)
        Dim arrData() As String: arrData = checkArr(1)
        Dim dictTarget As Object: Set dictTarget = checkArr(2)
        
        For i = LBound(arrData) To UBound(arrData)
            colName = Trim(arrData(i))
            If colName <> "" Then
                If Not dictHeaders.Exists(LCase(colName)) Then
                    CompareExcelRanges = Array(Array("Error: " & arrType & " column '" & colName & "' not found in headers."))
                    Exit Function
                End If
                If Not dictTarget.Exists(colName) Then
                    dictTarget.Add colName, dictHeaders(LCase(colName))
                End If
            End If
        Next i
    Next
    
    For Each compHeader In indexColIndexes.Keys
        If refColIndexes.Exists(compHeader) Then
            CompareExcelRanges = Array(Array("Error: Column '" & compHeader & "' is assigned to both Index and Reference categories."))
            Exit Function
        End If
    Next
    
    For i = 1 To colCount
        Dim currentHeader As String: currentHeader = Trim(CStr(header1(1, i)))
        Dim lowerHeader As String: lowerHeader = LCase(currentHeader)
        
        Dim isIndex As Boolean: isIndex = indexColIndexes.Exists(currentHeader)
        Dim isRef As Boolean: isRef = refColIndexes.Exists(currentHeader)
        Dim isIgnore As Boolean: isIgnore = False
        
        For Each colName In arrIgnore
            If StrComp(lowerHeader, LCase(Trim(CStr(colName))), vbTextCompare) = 0 Then
                isIgnore = True
                Exit For
            End If
        Next
        
        If Not isIndex And Not isRef And Not isIgnore Then
            compareColIndexes.Add currentHeader, i
        End If
    Next i
    
    ' --- 4. Build Dictionaries from Ranges ---
    
    Call PopulateDictionary(rng1, indexColIndexes, dictT1)
    Call PopulateDictionary(rng2, indexColIndexes, dictT2)

    ' --- 5. Comparison and Result Structuring ---
    
    ' Get all unique keys present in either table
    Dim allKeys As Object: Set allKeys = CreateObject("Scripting.Dictionary")
    
    ' FIX: Expanded the For Each loops to avoid "control variable already in use" error
    For Each key In dictT1.Keys
        If Not allKeys.Exists(key) Then allKeys.Add key, True
    Next key
    
    For Each key In dictT2.Keys
        If Not allKeys.Exists(key) Then allKeys.Add key, True
    Next key
    
    ' Iterate through all unique keys to perform comparison
    For Each key In allKeys.Keys
        Dim valuesT1 As Variant: valuesT1 = dictT1.Item(key)
        Dim valuesT2 As Variant: valuesT2 = dictT2.Item(key)
        
        Dim isT1Present As Boolean: isT1Present = Not IsEmpty(valuesT1)
        Dim isT2Present As Boolean: isT2Present = Not IsEmpty(valuesT2)
        
        Dim status As String
        Dim hasDiff As Boolean: hasDiff = False
        
        ' Check for difference (Mismatch) if both are present
        If isT1Present And isT2Present Then
            For Each compHeader In compareColIndexes.Keys
                i = compareColIndexes.Item(compHeader)
                If CStr(valuesT1(i)) <> CStr(valuesT2(i)) Then
                    hasDiff = True
                    Exit For
                End If
            Next
        Else ' If one is missing, it is a difference (Unique)
            hasDiff = True
        End If

        ' --- 1. SET NEW STATUS LABELS ---
        If isT1Present And isT2Present Then
            status = "both" ' ƒÉ‚€±í¶¼ÓÐ (both)
        ElseIf isT1Present Then
            status = "left_only" ' ±í1ÓÐ±í¶þŸo (left_only)
        ElseIf isT2Present Then
            status = "right_only" ' ±í¶þÓÐ±í1Ÿo (right_only)
        End If
        
        ' Skip if present in both and no difference ("No Change")
        If Not hasDiff And (isT1Present And isT2Present) Then GoTo NextKey

        Dim rowResult As Object: Set rowResult = CreateObject("Scripting.Dictionary")
        
        ' ----------------------------------------------------
        ' --- 2. POPULATE DICTIONARY IN DESIRED COLUMN ORDER ---
        '       Status -> Index -> Ref -> T1 -> T2 -> Diff
        ' ----------------------------------------------------
        
        ' --- 2a. Status Column ---
        rowResult.Add "Status", status
        
        ' --- 2b. Index Columns ---
        For Each compHeader In indexColIndexes.Keys
            i = indexColIndexes.Item(compHeader)
            If isT1Present Then
                rowResult.Add compHeader, valuesT1(i)
            ElseIf isT2Present Then
                rowResult.Add compHeader, valuesT2(i)
            End If
        Next
        
        ' --- 2c. Reference Columns (Ref Col) ---
        For Each compHeader In refColIndexes.Keys
            i = refColIndexes.Item(compHeader)
            Dim refVal As Variant: refVal = ""
            
            ' Fetch values safely
            Dim refT2 As Variant
            If isT2Present Then refT2 = valuesT2(i) Else refT2 = Empty
            Dim refT1 As Variant
            If isT1Present Then refT1 = valuesT1(i) Else refT1 = Empty
            
            ' VLOOKUP Logic
            If vlookupOrder Then ' TRUE: Priority T2 then T1
                If Not IsEmpty(refT2) And refT2 <> "" Then refVal = refT2
                If IsEmpty(refVal) Or refVal = "" Then
                    If Not IsEmpty(refT1) And refT1 <> "" Then refVal = refT1
                End If
            Else ' FALSE: Priority T1 then T2
                If Not IsEmpty(refT1) And refT1 <> "" Then refVal = refT1
                If IsEmpty(refVal) Or refVal = "" Then
                    If Not IsEmpty(refT2) And refT2 <> "" Then refVal = refT2
                End If
            End If
            
            rowResult.Add compHeader & "_Ref", refVal
        Next
        
        ' --- 2d. Prepare Comparison Columns for Grouping ---
        Dim t1Values As Object: Set t1Values = CreateObject("Scripting.Dictionary")
        Dim t2Values As Object: Set t2Values = CreateObject("Scripting.Dictionary")
        Dim diffValues As Object: Set diffValues = CreateObject("Scripting.Dictionary")
        
        For Each compHeader In compareColIndexes.Keys
            i = compareColIndexes.Item(compHeader)
            
            Dim val1 As Variant
            Dim val2 As Variant

            ' Safe retrieval logic (FIX from IIf)
            If isT1Present Then val1 = valuesT1(i) Else val1 = ""
            If isT2Present Then val2 = valuesT2(i) Else val2 = ""
            
            Dim diffVal As Variant: diffVal = ""
            
            If IsNumeric(val1) And IsNumeric(val2) Then
                diffVal = CDbl(val2) - CDbl(val1)
            ElseIf isT1Present And Not isT2Present Then
                If IsNumeric(val1) Then diffVal = 0 - CDbl(val1)
            ElseIf isT2Present And Not isT1Present Then
                If IsNumeric(val2) Then diffVal = CDbl(val2) - 0
            End If
            
            ' Store in temp Dictionaries
            t1Values.Add compHeader, val1
            t2Values.Add compHeader, val2
            diffValues.Add compHeader, diffVal
        Next
        
        ' --- 2e. Populate Comparison Columns into rowResult (Grouped Order) ---
        
        ' Group 1: Table 1 (T1)
        For Each compHeader In t1Values.Keys
            rowResult.Add compHeader & "_T1", t1Values.Item(compHeader)
        Next
        
        ' Group 2: Table 2 (T2)
        For Each compHeader In t2Values.Keys
            rowResult.Add compHeader & "_T2", t2Values.Item(compHeader)
        Next
        
        ' Group 3: Difference (Diff)
        For Each compHeader In diffValues.Keys
            rowResult.Add compHeader & "_Diff", diffValues.Item(compHeader)
        Next
        
        resultCollection.Add rowResult.Items
NextKey:
    Next key
    
    ' --- 6. Assemble Final 2D Array (Including Headers) ---
    
    Dim totalOutputCols As Long
    totalOutputCols = 1 + indexColIndexes.Count + refColIndexes.Count + (compareColIndexes.Count * 3)
    
    If resultCollection.Count = 0 Then
        CompareExcelRanges = Array(Array("Success: No differences found."))
        Exit Function
    End If
    
    Dim arrResult() As Variant
    ReDim arrResult(1 To resultCollection.Count + 2, 1 To totalOutputCols)
    
    ' --- 6a. Assemble Header Row A (Top Level: table1, table2, diff) ---
    Dim outputRow As Long: outputRow = 1
    Dim colIndex As Long: colIndex = 1
    
    ' Status, Index Cols, Ref Cols (Top row is blank/Ref for these)
    arrResult(outputRow, colIndex) = "": colIndex = colIndex + 1
    For i = 1 To indexColIndexes.Count: arrResult(outputRow, colIndex) = "": colIndex = colIndex + 1: Next
    For i = 1 To refColIndexes.Count: arrResult(outputRow, colIndex) = "Ref": colIndex = colIndex + 1: Next
    
    ' Group 1: T1 Comparison Columns
    For i = 1 To compareColIndexes.Count
        arrResult(outputRow, colIndex) = "table1"
        colIndex = colIndex + 1
    Next
    
    ' Group 2: T2 Comparison Columns
    For i = 1 To compareColIndexes.Count
        arrResult(outputRow, colIndex) = "table2"
        colIndex = colIndex + 1
    Next
    
    ' Group 3: Diff Comparison Columns
    For i = 1 To compareColIndexes.Count
        arrResult(outputRow, colIndex) = "diff"
        colIndex = colIndex + 1
    Next
    
    ' --- 6b. Assemble Header Row B (Column Names) ---
    outputRow = 2
    colIndex = 1
    arrResult(outputRow, colIndex) = "Status"
    colIndex = colIndex + 1

    For Each colName In indexColIndexes.Keys
        arrResult(outputRow, colIndex) = colName
        colIndex = colIndex + 1
    Next
    
    For Each colName In refColIndexes.Keys
        arrResult(outputRow, colIndex) = colName
        colIndex = colIndex + 1
    Next

    ' Group 1: T1 Column Names (Compare Cols)
    For Each compHeader In compareColIndexes.Keys
        arrResult(outputRow, colIndex) = compHeader
        colIndex = colIndex + 1
    Next
    
    ' Group 2: T2 Column Names (Compare Cols)
    For Each compHeader In compareColIndexes.Keys
        arrResult(outputRow, colIndex) = compHeader
        colIndex = colIndex + 1
    Next
    
    ' Group 3: Diff Column Names (Compare Cols)
    For Each compHeader In compareColIndexes.Keys
        arrResult(outputRow, colIndex) = compHeader
        colIndex = colIndex + 1
    Next

    ' --- 6c. Populate Data Rows ---
    outputRow = 3
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

' --- Updated Helper Subroutine to Handle Duplicates via Summation ---

Private Sub PopulateDictionary(ByVal rng As Range, ByVal indexCols As Object, ByRef targetDict As Object)
    ' Extracts data from a range (excluding header).
    ' [NEW LOGIC]: If a duplicate Index Key is found, numeric values are SUMMED.
    
    Dim dataRange As Range
    Dim dataArray As Variant
    Dim r As Long
    Dim c As Long
    Dim colIndex As Variant
    Dim keyString As String
    Dim existingValues As Variant
    Dim newValues() As Variant
    Dim arrayColCount As Long
    
    ' Adjust range to exclude the header row
    Set dataRange = rng.Offset(1, 0).Resize(rng.Rows.Count - 1, rng.Columns.Count)
    
    If dataRange.Rows.Count = 0 Then Exit Sub
    
    ' Read data into a single array for faster processing
    dataArray = dataRange.Value
    arrayColCount = UBound(dataArray, 2)
    
    ' Loop through every row in the source data
    For r = 1 To UBound(dataArray, 1)
        keyString = ""
        
        ' Build the key string
        For Each colIndex In indexCols.Items
            keyString = keyString & CStr(dataArray(r, colIndex)) & "|"
        Next colIndex
        
        ' Remove trailing separator
        If Len(keyString) > 0 Then keyString = Left(keyString, Len(keyString) - 1)
        
        ' Check if key already exists
        If Not targetDict.Exists(keyString) Then
            ' --- CASE 1: NEW KEY ---
            ' Create new array and store it
            ReDim newValues(1 To arrayColCount)
            For c = 1 To arrayColCount
                newValues(c) = dataArray(r, c)
            Next c
            targetDict.Add keyString, newValues
        Else
            ' --- CASE 2: DUPLICATE KEY (SUMMATION LOGIC) ---
            ' Retrieve existing data array
            existingValues = targetDict.Item(keyString)
            
            For c = 1 To arrayColCount
                ' Check if this column is part of the Index (Keys shouldn't be summed)
                ' We iterate through indexCols items to see if 'c' is an index column.
                Dim isIndexCol As Boolean
                isIndexCol = False
                
                For Each colIndex In indexCols.Items
                    If CLng(colIndex) = c Then
                        isIndexCol = True
                        Exit For
                    End If
                Next colIndex
                
                ' Only sum if it's NOT an index column AND both values are numeric
                If Not isIndexCol Then
                    If IsNumeric(existingValues(c)) And IsNumeric(dataArray(r, c)) Then
                        ' Perform Summation
                        existingValues(c) = CDbl(existingValues(c)) + CDbl(dataArray(r, c))
                    Else
                        ' Optional: If non-numeric data differs, you might want to mark it?
                        ' For now, we keep the FIRST value found (standard aggregation behavior for non-numerics)
                        ' or you could overwrite it: existingValues(c) = dataArray(r, c)
                    End If
                End If
            Next c
            
            ' Update the dictionary with the summed array
            targetDict.Item(keyString) = existingValues
        End If
    Next r
End Sub

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
    listCount = resultList.Count

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


' Purpose: Melts a table (wide-to-long format transformation), similar to pandas.melt().
' It preserves index columns and transforms other columns into variable-value pairs.
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
    
    ' Validate that the tableRange and idColumnsRange are aligned in height.
    If tableRange.Rows.Count <> idColumnsRange.Rows.Count Then
        MeltData = CVErr(xlErrValue) ' Return #VALUE! error
        Exit Function
    End If

    ' Get ID column count from the passed range.
    idColCount = idColumnsRange.Columns.Count
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
            
            currentResultRow = currentResultRow + 1
        Next dataCol
    Next dataRow
    
    ' --- Return the Array ---
    MeltData = result
    
End Function

