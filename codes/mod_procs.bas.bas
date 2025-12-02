Attribute VB_Name = "mod_procs"

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
