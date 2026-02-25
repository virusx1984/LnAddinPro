Option Explicit

' ==============================================================================
' SUB: Manual_Register_LNF
' Purpose: A wrapper to safely register function descriptions.
'          It checks if Excel is ready (has visible windows) to avoid Error 1004.
'          It also saves the Add-in to persist changes.
' ==============================================================================
Public Sub Manual_Register_LNF()
    
    ' 1. Check if there are any visible workbooks open.
    '    Application.MacroOptions requires a visible window context to run safely.
    If Application.Windows.count = 0 Then
        MsgBox "Cannot register functions when no workbook is open." & vbCrLf & _
               "Please open or create a blank workbook and try again.", _
               vbExclamation, "Registration Skipped"
        Exit Sub
    End If
    
    ' 2. Error Handling to catch unexpected issues
    On Error GoTo ErrHandler
    
    ' 3. Call the main registration routine
    Call Register_LNF_Functions
    
    ' 4. Critical: Save the Add-in itself to persist the descriptions!
    '    If you don't save ThisWorkbook, the descriptions will be lost on restart.
    If ThisWorkbook.IsAddin Then
        ThisWorkbook.Save
    End If
    
    ' 5. Success Message
    MsgBox "LNF Functions registered and Add-in saved successfully!", _
           vbInformation, "Success"
    
    Exit Sub

ErrHandler:
    MsgBox "An error occurred during registration: " & Err.Description, _
           vbCritical, "Error " & Err.Number
End Sub

' ==============================================================================
' SUB: Register_LNF_Functions
' Purpose: Registers function descriptions and argument help text for the
'          Excel Function Wizard (Shift + F3).
' Usage:   Call this from Workbook_Open in ThisWorkbook.
' ==============================================================================
Public Sub Register_LNF_Functions()
    
    Dim categoryName As String
    categoryName = "LNF Tools" ' This category will appear in the Function Wizard
    
    ' --- 1. LNF_Join ---
    Application.MacroOptions _
        Macro:="LNF_Join", _
        Description:="Concatenates text from a range into a single string with optional delimiters and surrounding characters.", _
        category:=categoryName, _
        ArgumentDescriptions:=Array( _
            "sourceRange: The range of cells to join.", _
            "joinString: The delimiter string (e.g., comma).", _
            "leftSurround: (Optional) Character(s) to add before each item.", _
            "rightSurround: (Optional) Character(s) to add after each item. Defaults to leftSurround if omitted." _
        )

    ' --- 2. LNF_RegexExtract ---
    Application.MacroOptions _
        Macro:="LNF_RegexExtract", _
        Description:="Extracts the first substring matching a Regular Expression pattern.", _
        category:=categoryName, _
        ArgumentDescriptions:=Array( _
            "sourceText: The original text to search.", _
            "pattern: The Regex pattern (e.g., ""\d+"" for numbers).", _
            "ignoreCase: (Optional) True to ignore case. Default is True." _
        )

    ' --- 3. LNF_ExtractNumber ---
    Application.MacroOptions _
        Macro:="LNF_ExtractNumber", _
        Description:="Removes non-numeric characters from a string, keeping only digits, decimal points, and leading negative signs.", _
        category:=categoryName, _
        ArgumentDescriptions:=Array( _
            "sourceText: The dirty string containing numbers (e.g., ""USD 1,200.00"")." _
        )

    ' --- 4. LNF_GetLastRow ---
    Application.MacroOptions _
        Macro:="LNF_GetLastRow", _
        Description:="Returns the row number of the last non-empty cell in a specific column.", _
        category:=categoryName, _
        ArgumentDescriptions:=Array( _
            "ws: The target Worksheet object.", _
            "col: (Optional) The column index (1) or letter (""A""). Default is 1." _
        )
        
    ' --- 5. LNF_Exists ---
    Application.MacroOptions _
        Macro:="LNF_Exists", _
        Description:="Checks if a specific value exists within a Range or an Array.", _
        category:=categoryName, _
        ArgumentDescriptions:=Array( _
            "valueToFind: The value to search for.", _
            "sourceContainer: The search scope (Range object or Array)." _
        )

    ' --- 6. LNF_VLookupNth ---
    Application.MacroOptions _
        Macro:="LNF_VLookupNth", _
        Description:="Advanced VLookup that retrieves the N-th match instead of just the first one.", _
        category:=categoryName, _
        ArgumentDescriptions:=Array( _
            "lookupVal: The value to look up.", _
            "searchRng: The range to search within (e.g., Column A).", _
            "returnColOffset: The number of columns to the right to retrieve data from (0 = same column).", _
            "matchIndex: Which match instance to return (1 = first, 2 = second, etc.)." _
        )
    
    ' Optional: Confirmation message for debugging (comment out in production)
    ' MsgBox "LNF Function descriptions registered successfully.", vbInformation
    
End Sub

' ==============================================================================
' FUNCTION: LNF_Join
' Purpose:  Concatenates text from a given range of cells.
' ==============================================================================
Public Function LNF_Join(ByVal sourceRange As Range, ByVal joinString As String, _
                         Optional ByVal leftSurround As String = "", _
                         Optional ByVal rightSurround As String = "") As String
    
    Dim cell As Range
    Dim resultString As String
    
    ' If the right surround string is not provided, use the left surround string.
    If rightSurround = "" Then rightSurround = leftSurround
    
    ' Loop through each cell in the provided range.
    For Each cell In sourceRange
        ' Avoid processing empty cells unless specifically desired.
        ' This prevents extra delimiters in the final string.
        If Not IsEmpty(cell.Value) And Len(CStr(cell.Value)) > 0 Then
            ' Append the join string if the result string is not empty.
            If Len(resultString) > 0 Then
                resultString = resultString & joinString
            End If
            
            ' Append the surrounding characters and the cell's text.
            resultString = resultString & leftSurround & CStr(cell.Value) & rightSurround
        End If
    Next cell
    
    LNF_Join = resultString
    
End Function

' ==============================================================================
' FUNCTION: LNF_RegexExtract
' Purpose:  Extracts a substring using Regular Expressions.
' Note:     Uses Late Binding (CreateObject) to avoid reference issues.
' ==============================================================================
Public Function LNF_RegexExtract(ByVal sourceText As String, ByVal pattern As String, _
                                 Optional ByVal ignoreCase As Boolean = True) As String
    Dim regEx As Object
    Dim matches As Object
    
    On Error Resume Next
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Global = False      ' Return only the first match
        .MultiLine = False
        .ignoreCase = ignoreCase
        .pattern = pattern
    End With
    
    If regEx.Test(sourceText) Then
        Set matches = regEx.Execute(sourceText)
        LNF_RegexExtract = matches(0).Value
    Else
        LNF_RegexExtract = ""
    End If
    On Error GoTo 0
End Function

' ==============================================================================
' FUNCTION: LNF_ExtractNumber
' Purpose:  Cleans a string to return only numeric values.
'           Handles decimal points and leading negative signs.
' ==============================================================================
Public Function LNF_ExtractNumber(ByVal sourceText As String) As Double
    Dim i As Integer
    Dim strResult As String
    Dim char As String
    Dim hasDecimal As Boolean
    
    strResult = ""
    hasDecimal = False
    
    For i = 1 To Len(sourceText)
        char = Mid(sourceText, i, 1)
        
        ' Allow digits
        If IsNumeric(char) Then
            strResult = strResult & char
        
        ' Allow one decimal point
        ElseIf char = "." And Not hasDecimal Then
            strResult = strResult & char
            hasDecimal = True
            
        ' Allow negative sign only at the very beginning
        ElseIf char = "-" And Len(strResult) = 0 Then
            strResult = strResult & char
        End If
    Next i
    
    If Len(strResult) > 0 Then
        LNF_ExtractNumber = Val(strResult)
    Else
        LNF_ExtractNumber = 0
    End If
End Function

' ==============================================================================
' FUNCTION: LNF_GetLastRow
' Purpose:  Finds the last used row in a specific column.
'           More reliable than UsedRange.
' ==============================================================================
Public Function LNF_GetLastRow(ByVal ws As Worksheet, Optional ByVal col As Variant = 1) As Long
    On Error Resume Next
    ' Equivalent to Ctrl+Up from the bottom of the sheet
    LNF_GetLastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
    On Error GoTo 0
End Function

' ==============================================================================
' FUNCTION: LNF_Exists
' Purpose:  Checks if a value exists in a Range or Array.
' ==============================================================================
Public Function LNF_Exists(ByVal valueToFind As Variant, ByVal sourceContainer As Variant) As Boolean
    Dim item As Variant
    
    LNF_Exists = False
    
    ' Case 1: Container is a Range object
    If TypeName(sourceContainer) = "Range" Then
        Dim rng As Range
        ' LookAt:=xlWhole ensures exact match
        Set rng = sourceContainer.Find(What:=valueToFind, LookIn:=xlValues, _
                                       LookAt:=xlWhole, MatchCase:=False)
        If Not rng Is Nothing Then LNF_Exists = True
        
    ' Case 2: Container is an Array
    ElseIf IsArray(sourceContainer) Then
        For Each item In sourceContainer
            If CStr(item) = CStr(valueToFind) Then
                LNF_Exists = True
                Exit Function
            End If
        Next item
    End If
End Function

' ==============================================================================
' FUNCTION: LNF_VLookupNth
' Purpose:  Performs a lookup but retrieves the N-th match.
' ==============================================================================
Public Function LNF_VLookupNth(ByVal lookupVal As Variant, ByVal searchRng As Range, _
                               ByVal returnColOffset As Integer, ByVal matchIndex As Integer) As Variant
    Dim cell As Range
    Dim count As Integer
    
    count = 0
    LNF_VLookupNth = Empty ' Return empty if not found
    
    For Each cell In searchRng
        If cell.Value = lookupVal Then
            count = count + 1
            If count = matchIndex Then
                LNF_VLookupNth = cell.offset(0, returnColOffset).Value
                Exit Function
            End If
        End If
    Next cell
End Function