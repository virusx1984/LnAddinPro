' ==============================================================================
' Purpose: Scans the active sheet for External links, LNF_ functions, and Cross-Sheet references.
' ==============================================================================
' Configuration: Max number of items to collect for the UI ListBox
' Prevents UI freeze when thousands of errors exist.
Private Const UI_DISPLAY_LIMIT As Long = 200

Public Sub LNS_ShowLinkChecker(control As IRibbonControl)
    Dim rngFormulas As Range
    Dim area As Range
    Dim vFormulas As Variant
    Dim arrResults() As Variant
    
    Dim r As Long, c As Long
    Dim i As Long
    Dim fmla As String
    Dim cellAddr As String
    
    Dim count As Long           ' Items collected for display
    Dim totalFound As Long      ' Total items found (statistics)
    
    Dim isExt As Boolean, isLNF As Boolean, isInt As Boolean
    
    ' 1. Fast Check: Get Formula Cells
    On Error Resume Next
    Set rngFormulas = ActiveSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    
    If rngFormulas Is Nothing Then
        MsgBox "No formulas found in the active sheet.", vbInformation, "Link Checker"
        Exit Sub
    End If
    
    ' 2. Initialize Output Array (Fixed size for speed, trimmed later)
    ' Format: 1=Address, 2=Type, 3=Formula
    ReDim arrResults(1 To UI_DISPLAY_LIMIT, 1 To 3)
    count = 0
    totalFound = 0
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' Pause Calc for speed
    
    ' 3. Iterate through Areas (Batch Processing)
    For Each area In rngFormulas.Areas
        
        ' Case A: Single Cell Area
        If area.Cells.count = 1 Then
            fmla = area.Formula
            If CheckFormula(fmla, isExt, isLNF, isInt) Then
                totalFound = totalFound + 1
                If count < UI_DISPLAY_LIMIT Then
                    count = count + 1
                    StoreResult arrResults, count, area.Address(False, False), fmla, isExt, isLNF, isInt
                End If
            End If
            
        ' Case B: Multi-Cell Area (Read into Array)
        Else
            vFormulas = area.Formula ' Copy block to memory (Fastest)
            
            ' vFormulas is always 2D array (1 to Rows, 1 to Cols)
            For r = 1 To UBound(vFormulas, 1)
                For c = 1 To UBound(vFormulas, 2)
                    fmla = vFormulas(r, c)
                    
                    ' Processing logic in memory
                    If CheckFormula(fmla, isExt, isLNF, isInt) Then
                        totalFound = totalFound + 1
                        
                        ' Only store if under limit
                        If count < UI_DISPLAY_LIMIT Then
                            count = count + 1
                            ' Calculate Address manually is slow, use .Cells(r,c).Address only when hit
                            ' Optimization: Relative to Area
                            cellAddr = area.Cells(r, c).Address(False, False)
                            StoreResult arrResults, count, cellAddr, fmla, isExt, isLNF, isInt
                        End If
                    End If
                    
                Next c
            Next r
        End If
        
    Next area
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' 4. Results Handling
    If totalFound = 0 Then
        MsgBox "Clean Sheet! No External, LNF, or Cross-Sheet links found.", vbInformation, "Link Checker"
    Else
        ' Trim array to actual stored count
        Dim finalData() As Variant
        Dim j As Long, k As Long
        
        If count > 0 Then
            ReDim finalData(1 To count, 1 To 3)
            For j = 1 To count
                For k = 1 To 3
                    finalData(j, k) = arrResults(j, k)
                Next k
            Next j
        End If
        
        ' Initialize Form with Data AND Statistics
        ' Ensure UserForm is named "frm_link_checker"
        Dim frm As New frm_link_checker
        frm.LoadData finalData, totalFound
        frm.Show vbModeless
    End If

End Sub

' Helper: Centralized Logic check
Private Function CheckFormula(ByVal f As String, ByRef bExt As Boolean, ByRef bLNF As Boolean, ByRef bInt As Boolean) As Boolean
    bExt = False: bLNF = False: bInt = False
    
    ' Priority 1: External ([)
    If InStr(1, f, "[", vbBinaryCompare) > 0 Then
        bExt = True
        CheckFormula = True
        Exit Function
    End If
    
    ' Priority 2: LNF Custom Func
    If InStr(1, f, "LNF_", vbTextCompare) > 0 Then
        bLNF = True
        CheckFormula = True
        Exit Function
    End If
    
    ' Priority 3: Internal (!)
    If InStr(1, f, "!", vbBinaryCompare) > 0 Then
        bInt = True
        CheckFormula = True
        Exit Function
    End If
    
    CheckFormula = False
End Function

' Helper: Store data into array
Private Sub StoreResult(ByRef arr() As Variant, idx As Long, addr As String, fmla As String, bExt As Boolean, bLNF As Boolean, bInt As Boolean)
    arr(idx, 1) = addr
    arr(idx, 3) = fmla
    
    If bExt Then
        arr(idx, 2) = "External"
    ElseIf bLNF Then
        arr(idx, 2) = "LNF_Func"
    Else
        arr(idx, 2) = "Internal"
    End If
End Sub


' ==============================================================================
' Purpose: Entry point for the Format Table Ribbon button.
' ==============================================================================
Public Sub LNS_ShowFormatDialog(control As IRibbonControl)
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    ' Show the mini configuration form
    ' (Ensure userform is named 'frm_format_mini')
    frm_format_mini.Show
End Sub

' ==============================================================================
' Purpose: Launches the Duplicate Checker Wizard (frm_duplicate_check).
'          1. Checks if a workbook is active.
'          2. Checks if the selection is valid (Range).
'          3. Shows the UserForm.
' ==============================================================================
Public Sub LNS_ShowDuplicateChecker(control As IRibbonControl)
    
    ' 1. Check if a Workbook is open
    If ActiveWorkbook Is Nothing Then
        MsgBox "Please open a workbook first.", vbExclamation, "LnAddinPro"
        Exit Sub
    End If
    
    ' 2. Check if selection is a Range (Optional safety check matching your style)
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbExclamation, "LnAddinPro"
        Exit Sub
    End If
    
    ' 3. Load and Show the UserForm
    '    Note: The form logic (frm_duplicate_check) handles the specific
    '    limit checks (100k cells) upon initialization or analysis.
    Load frm_duplicate_check
    frm_duplicate_check.Show
    
End Sub


' ==============================================================================
' Purpose: Colors selected cells based on Boolean values.
'          True  -> Green (vbGreen)
'          False -> Red (vbRed)
'          Others -> Unchanged
' ==============================================================================
Public Sub LNS_ColorBooleanValues(control As IRibbonControl)
    Dim cell As Range
    Dim selectedRng As Range
    Dim cellVal As Variant
    
    ' 1. Check if selection is a Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first.", vbExclamation, "LnAddinPro"
        Exit Sub
    End If
    
    Set selectedRng = Selection
    
    ' Optimize performance
    Application.ScreenUpdating = False
    
    ' 2. Loop through each cell
    On Error Resume Next ' Prevent errors on special cell types
    For Each cell In selectedRng
        cellVal = cell.Value
        
        ' Skip Error cells and Empty cells
        If Not IsError(cellVal) And Not IsEmpty(cellVal) Then
            
            ' Check for TRUE (Boolean or Text)
            If UCase(CStr(cellVal)) = "TRUE" Then
                cell.Interior.Color = vbGreen
                
            ' Check for FALSE (Boolean or Text)
            ElseIf UCase(CStr(cellVal)) = "FALSE" Then
                cell.Interior.Color = vbRed
            End If
            
            ' Note: Other values remain unchanged
        End If
    Next cell
    On Error GoTo 0
    
    ' Restore
    Application.ScreenUpdating = True
End Sub

' ==============================================================================
' Purpose: Sets the standard format for the active sheet:
'          1. Hides Gridlines.
'          2. Hides Page Breaks (dashed lines).
'          3. Sets Font to Arial, Size 10.
'          4. Sets Vertical Alignment to Center.
' ==============================================================================
Public Sub LNS_ApplyStandardFormat(control As IRibbonControl)
    On Error Resume Next
    
    ' Optimize performance
    Application.ScreenUpdating = False
    
    ' 1. Hide Gridlines for the active window
    ActiveWindow.DisplayGridlines = False
    
    ' 2. Hide Page Breaks (dashed lines)
    ActiveSheet.DisplayPageBreaks = False
    
    ' 3. Apply Font settings to all cells
    With ActiveSheet.Cells.Font
        .Name = "Arial"
        .Size = 10
    End With
    
    ' 4. Set Vertical Alignment to Center for all cells
    ActiveSheet.Cells.VerticalAlignment = xlCenter
    
    ' Optional: Select A1 to reset selection
    ActiveSheet.Range("A1").Select
    
    ' Restore
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' ==============================================================================
' Purpose: Registers custom functions (UDFs) descriptions into Excel.
' The control parameter is REQUIRED by the Ribbon's onAction mechanism.
' ==============================================================================
Public Sub LNS_RegisterFunctions(control As IRibbonControl)
    ' LNS: LnAddinPro Sub - Manually triggers the function description registration.
    
    ' Call the wrapper function in mod_lnf
    ' Make sure Manual_Register_LNF is Public in mod_lnf
    Run "Manual_Register_LNF"
    
End Sub

' Purpose: Launches the About UserForm.
Public Sub LNS_ShowAboutForm(control As IRibbonControl)
    Load frm_about
    frm_about.Show
End Sub

' Purpose: Launch the UserForm
Public Sub LNS_ShowCodeExportForm(control As IRibbonControl)
    Load frm_code_export
    frm_code_export.Show
End Sub


' Purpose: Launches the JSON Export UserForm (frm_json_export).
' The control parameter is REQUIRED by the Ribbon's onAction mechanism.
Public Sub LNS_ShowJsonExportForm(control As IRibbonControl)
    ' LNS: LnAddinPro Sub - Shows the UserForm for exporting range to JSON.
    
    ' Load the form into memory
    Load frm_json_export
    
    ' Display the form
    frm_json_export.Show
End Sub


' Purpose: Launches the Data Melt UserForm (frm_melt).
' The control parameter is REQUIRED by the Ribbon's onAction mechanism.
Public Sub LNS_ShowMeltForm(control As IRibbonControl)
    ' LNS: LnAddinPro Sub - Shows the UserForm for melting data.
    
    ' Load the form into memory
    Load frm_melt
    
    ' Display the form
    frm_melt.Show
End Sub


' Purpose: Launches the Time Series Generator UserForm (frm_gen_time_series).
' The control parameter is REQUIRED by the Ribbon's onAction mechanism.
Public Sub LNS_ShowTimeSeriesForm(control As IRibbonControl)
    ' LNS: LnAddinPro Sub - Shows the UserForm for generating time series.

    ' Load the form into memory
    Load frm_gen_time_series
    
    ' Display the form
    frm_gen_time_series.Show
End Sub

' Purpose: Launches the Compare Setup UserForm (frm_compare_setup).
' The control parameter is REQUIRED by the Ribbon's onAction mechanism.
Public Sub LNS_ShowCompareForm(control As IRibbonControl)
    ' LNS: LnAddinPro Sub - Shows the UserForm for comparing data ranges.
    
    ' Load the form into memory
    Load frm_compare_setup
    
    ' Display the form
    frm_compare_setup.Show
End Sub