Attribute VB_Name = "mod_procs"
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
