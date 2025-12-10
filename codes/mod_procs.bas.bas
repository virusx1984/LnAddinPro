Attribute VB_Name = "mod_procs"
' ==============================================================================
' Purpose: Sets the standard format for the active sheet:
'          1. Hides Gridlines.
'          2. Sets Font to Arial.
'          3. Sets Font Size to 10.
' ==============================================================================
Public Sub LNS_ApplyStandardFormat(control As IRibbonControl)
    On Error Resume Next
    
    ' Optimize performance
    Application.ScreenUpdating = False
    
    ' 1. Hide Gridlines for the active window
    ActiveWindow.DisplayGridlines = False
    
    ' 2. Apply Font settings to all cells
    With ActiveSheet.Cells.Font
        .Name = ""
    End With
    With ActiveSheet.Cells.Font
        .Name = "Arial"
        .Size = 10
    End With
    
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
