Attribute VB_Name = "frm_gen_time_series"
Attribute VB_Base = "0{FB5933D1-02B0-442C-836A-977D87CFB6E5}{5FB3C614-FF53-459F-A5E2-BE9C716ADEB4}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Declare CommandButton variables using WithEvents to capture the Click event
' for dynamically created controls. This is essential for event routing.
Public WithEvents cmdExecute As MSForms.CommandButton
Attribute cmdExecute.VB_VarHelpID = -1
Public WithEvents cmdClose As MSForms.CommandButton
Attribute cmdClose.VB_VarHelpID = -1

' Declare input controls as Public Object for access within the Click event handler.
' Using Object ensures compatibility, especially for the RefEdit control.
Public txtIntervalType As Object
Public txtStartYear As Object
Public ckbIncludeAnnualTotal As Object
Public refOutputCell As Object


Private Sub UserForm_Initialize()
    
    ' --- 1. Form and Layout Setup ---
    Me.Caption = "Generate Time Series"
    Me.Width = 320
    Me.Height = 240
    
    ' Layout Constants
    Const LABEL_WIDTH As Long = 100
    Const CONTROL_WIDTH As Long = 180
    Const TOP_MARGIN As Long = 10
    Const SPACING As Long = 10
    Const CONTROL_HEIGHT As Long = 20
    Dim currentTop As Long: currentTop = TOP_MARGIN
    
    ' Calculated left position for input controls
    Dim controlLeft As Long: controlLeft = SPACING + LABEL_WIDTH + 5

    ' --- 2. Input Control Setup ---

    ' 2a. Interval Type (Label and TextBox)
    
    ' Label: lblIntervalType
    With Me.Controls.Add("Forms.Label.1", "lblIntervalType", True)
        .Caption = "1. Interval Combo:"
        .Width = LABEL_WIDTH
        .Height = CONTROL_HEIGHT
        .Left = SPACING
        .Top = currentTop + 3
    End With
    
    ' Control: txtIntervalType (Assigned to Public variable)
    Set txtIntervalType = Me.Controls.Add("Forms.TextBox.1", "txtIntervalType", True)
    With txtIntervalType
        .Left = controlLeft
        .Top = currentTop
        .Width = CONTROL_WIDTH
        .Height = CONTROL_HEIGHT
        .Value = "MMQHY"
    End With
    currentTop = currentTop + CONTROL_HEIGHT + SPACING

    ' 2b. Start Year (Label and TextBox)
    
    ' Label: lblStartYear
    With Me.Controls.Add("Forms.Label.1", "lblStartYear", True)
        .Caption = "2. Start Year (YYYY):"
        .Width = LABEL_WIDTH
        .Height = CONTROL_HEIGHT
        .Left = SPACING
        .Top = currentTop + 3
    End With
    
    ' Control: txtStartYear
    Set txtStartYear = Me.Controls.Add("Forms.TextBox.1", "txtStartYear", True)
    With txtStartYear
        .Left = controlLeft
        .Top = currentTop
        .Width = CONTROL_WIDTH / 2
        .Height = CONTROL_HEIGHT
        .Value = Year(Date)
    End With
    currentTop = currentTop + CONTROL_HEIGHT + SPACING

    ' 2c. Include Annual Total (CheckBox)
    
    ' Control: ckbIncludeAnnualTotal
    Set ckbIncludeAnnualTotal = Me.Controls.Add("Forms.CheckBox.1", "ckbIncludeAnnualTotal", True)
    With ckbIncludeAnnualTotal
        .Left = controlLeft
        .Top = currentTop
        .Width = CONTROL_WIDTH
        .Height = CONTROL_HEIGHT
        .Caption = "Include Annual Total (M, Q, H)"
        .Value = True
    End With
    currentTop = currentTop + CONTROL_HEIGHT + SPACING

    ' 2d. Output Cell (Label and RefEdit)
    
    ' Label: lblOutputCell
    With Me.Controls.Add("Forms.Label.1", "lblOutputCell", True)
        .Caption = "3. Output Cell:"
        .Width = LABEL_WIDTH
        .Height = CONTROL_HEIGHT
        .Left = SPACING
        .Top = currentTop + 3
    End With
    
    ' Control: refOutputCell (Using the correct RefEdit ProgID and assigning to Public variable)
    Set refOutputCell = Me.Controls.Add("RefEdit.Ctrl", "refOutputCell", True)
    With refOutputCell
        .Left = controlLeft
        .Top = currentTop
        .Width = CONTROL_WIDTH
        .Height = CONTROL_HEIGHT
        ' Set the RefEdit to the active cell initially
        If TypeName(Selection) = "Range" Then
            .Text = Selection.Address(External:=False)
        End If
    End With
    currentTop = currentTop + CONTROL_HEIGHT + SPACING + 15 ' Extra space before buttons

    ' --- 3. Command Buttons (Bound via WithEvents) ---
    
    Const BUTTON_WIDTH As Long = 80
    Dim buttonLeft As Long: buttonLeft = Me.Width - BUTTON_WIDTH - 70
    
    ' 3a. cmdExecute (Assigned to Public WithEvents cmdExecute variable)
    Set cmdExecute = Me.Controls.Add("Forms.CommandButton.1", "cmdExecute", True)
    With cmdExecute
        .Left = buttonLeft
        .Top = currentTop
        .Width = BUTTON_WIDTH
        .Height = 25
        .Caption = "Execute"
    End With
    currentTop = currentTop + 25 + SPACING

    ' 3b. cmdClose (Assigned to Public WithEvents cmdClose variable)
    Set cmdClose = Me.Controls.Add("Forms.CommandButton.1", "cmdClose", True)
    With cmdClose
        .Left = buttonLeft
        .Top = currentTop
        .Width = BUTTON_WIDTH
        .Height = 25
        .Caption = "Close"
    End With
    currentTop = currentTop + 25
    
    ' --- 4. Final Form Resize ---
    Me.Height = currentTop + TOP_MARGIN

End Sub


' Event handler triggered by the cmdExecute variable (declared WithEvents).
Private Sub cmdExecute_Click()
    Dim intervalType As String
    Dim includeTotal As Boolean
    Dim StartYear As Long
    Dim outputRange As Range
    Dim timeSeriesData As Variant
    Dim seriesLength As Long
    Dim ws As Worksheet
    
    ' 1. Validate and retrieve parameters (Accessing controls via their Public variable names)
    
    intervalType = Trim(txtIntervalType.Value) ' Using txtIntervalType variable
    If intervalType = "" Then
        MsgBox "Please enter the Interval Type Combo (e.g., MMQHY).", vbExclamation, "Input Error"
        txtIntervalType.SetFocus
        Exit Sub
    End If
    
    includeTotal = ckbIncludeAnnualTotal.Value ' Using ckbIncludeAnnualTotal variable
    
    ' Validate Start Year input
    If Not IsNumeric(txtStartYear.Value) Then ' Using txtStartYear variable
        MsgBox "Please enter a valid Start Year.", vbExclamation, "Input Error"
        txtStartYear.SetFocus
        Exit Sub
    End If
    StartYear = CLng(txtStartYear.Value)

    ' 2. Validate output cell
    On Error Resume Next
    ' Accessing RefEdit control via its Public variable name
    Set outputRange = Application.Range(refOutputCell.Text)
    On Error GoTo 0
    
    If outputRange Is Nothing Then
        MsgBox "Invalid output cell reference.", vbExclamation, "Reference Error"
        refOutputCell.SetFocus
        Exit Sub
    End If
    
    ' Use only the top-left cell of the selection
    Set outputRange = outputRange.Cells(1, 1)
    Set ws = outputRange.Worksheet

    ' 3. Generate Time Series Data
    ' Explicitly call the function from the mod_funcs module
    timeSeriesData = mod_funcs.GenerateTimeSeries(intervalType, includeTotal, StartYear)
    
    ' 4. Output the results
    If IsArray(timeSeriesData) Then
        ' Determine the number of elements (columns) to output
        seriesLength = UBound(timeSeriesData)
        
        ' Ensure the target sheet is active before writing (UX improvement)
        ws.Activate
        
        ' Output the 1D array horizontally to the resized range
        outputRange.Resize(1, seriesLength).Value = timeSeriesData
        
        MsgBox "Time series generated successfully! (" & seriesLength & " elements) to " & outputRange.Address(External:=False), vbInformation, "Success"
        
        ' Close the form after successful execution
        Unload Me
    Else
        MsgBox "Failed to generate time series or result is empty.", vbCritical, "Execution Error"
    End If
    
End Sub

' Event handler triggered by the cmdClose variable (declared WithEvents).
Private Sub cmdClose_Click()
    ' Close the form without performing any action
    Unload Me
End Sub
