Attribute VB_Name = "frm_gen_time_series"
Attribute VB_Base = "0{CD0F6040-69F0-40CA-B6B1-B8915EDCC19F}{86DF188E-61CD-4039-B28D-A179EEA0EEDF}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Declare CommandButton variables using WithEvents to capture the Click event
Public WithEvents cmdExecute As MSForms.CommandButton
Attribute cmdExecute.VB_VarHelpID = -1
Public WithEvents cmdClose As MSForms.CommandButton
Attribute cmdClose.VB_VarHelpID = -1

' Declare input controls as Public Object
Public txtIntervalType As Object
Public txtStartYear As Object
Public ckbIncludeAnnualTotal As Object
Public refOutputCell As Object

Private Sub UserForm_Initialize()
    
    ' --- 1. Form and Layout Setup ---
    Me.Caption = "Generate Time Series"
    ' Set width first. Note: We will set Height at the very end.
    Me.Width = 320
    
    ' Layout Constants
    Const LABEL_WIDTH As Long = 100
    Const CONTROL_WIDTH As Long = 180
    Const TOP_MARGIN As Long = 10
    Const SPACING As Long = 10
    Const CONTROL_HEIGHT As Long = 20
    
    Dim currentTop As Long: currentTop = TOP_MARGIN
    Dim controlLeft As Long: controlLeft = SPACING + LABEL_WIDTH + 5

    ' --- 2. Input Control Setup ---

    ' 2a. Interval Type
    With Me.Controls.Add("Forms.Label.1", "lblIntervalType", True)
        .Caption = "1. Interval Combo:"
        .Width = LABEL_WIDTH
        .Height = CONTROL_HEIGHT
        .Left = SPACING
        .Top = currentTop + 3
    End With
    
    Set txtIntervalType = Me.Controls.Add("Forms.TextBox.1", "txtIntervalType", True)
    With txtIntervalType
        .Left = controlLeft
        .Top = currentTop
        .Width = CONTROL_WIDTH
        .Height = CONTROL_HEIGHT
        .Value = "MMQHY"
    End With
    currentTop = currentTop + CONTROL_HEIGHT + SPACING

    ' 2b. Start Year
    With Me.Controls.Add("Forms.Label.1", "lblStartYear", True)
        .Caption = "2. Start Year (YYYY):"
        .Width = LABEL_WIDTH
        .Height = CONTROL_HEIGHT
        .Left = SPACING
        .Top = currentTop + 3
    End With
    
    Set txtStartYear = Me.Controls.Add("Forms.TextBox.1", "txtStartYear", True)
    With txtStartYear
        .Left = controlLeft
        .Top = currentTop
        .Width = CONTROL_WIDTH / 2
        .Height = CONTROL_HEIGHT
        .Value = Year(Date)
    End With
    currentTop = currentTop + CONTROL_HEIGHT + SPACING

    ' 2c. Include Annual Total
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

    ' 2d. Output Cell
    With Me.Controls.Add("Forms.Label.1", "lblOutputCell", True)
        .Caption = "3. Output Cell:"
        .Width = LABEL_WIDTH
        .Height = CONTROL_HEIGHT
        .Left = SPACING
        .Top = currentTop + 3
    End With
    
    Set refOutputCell = Me.Controls.Add("RefEdit.Ctrl", "refOutputCell", True)
    With refOutputCell
        .Left = controlLeft
        .Top = currentTop
        .Width = CONTROL_WIDTH
        .Height = CONTROL_HEIGHT
        If TypeName(Selection) = "Range" Then
            .Text = Selection.Address(External:=False)
        End If
    End With
    
    ' Add extra padding before the buttons
    currentTop = currentTop + CONTROL_HEIGHT + 20

    ' --- 3. Command Buttons ---
    
    Const BUTTON_WIDTH As Long = 80
    Const BUTTON_HEIGHT As Long = 24
    
    ' Calculate Right Alignment for buttons
    ' Me.InsideWidth ensures we calculate based on the usable area
    Dim btnLeftExecute As Long
    btnLeftExecute = Me.InsideWidth - BUTTON_WIDTH - SPACING
    
    ' 3a. cmdExecute
    Set cmdExecute = Me.Controls.Add("Forms.CommandButton.1", "cmdExecute", True)
    With cmdExecute
        .Left = btnLeftExecute
        .Top = currentTop
        .Width = BUTTON_WIDTH
        .Height = BUTTON_HEIGHT
        .Caption = "Execute"
        .Default = True ' Allows pressing Enter key
    End With
    
    ' Move currentTop down for the next button
    currentTop = currentTop + BUTTON_HEIGHT + SPACING

    ' 3b. cmdClose
    Set cmdClose = Me.Controls.Add("Forms.CommandButton.1", "cmdClose", True)
    With cmdClose
        .Left = btnLeftExecute
        .Top = currentTop
        .Width = BUTTON_WIDTH
        .Height = BUTTON_HEIGHT
        .Caption = "Close"
        .Cancel = True ' Allows pressing Esc key
    End With
    
    ' Increment currentTop to include the last button's height
    currentTop = currentTop + BUTTON_HEIGHT

    ' --- 4. Final Form Resize (CORRECTED) ---
    
    Dim formChromeHeight As Single
    
    ' Calculate the height of the Title Bar + Borders
    ' We do this by subtracting the usable area (InsideHeight) from the total area (Height)
    formChromeHeight = Me.Height - Me.InsideHeight
    
    ' Now set the total Height.
    ' Logic: Total Height = (Where the last control ends) + (Bottom Margin) + (TitleBar/Borders)
    Me.Height = currentTop + TOP_MARGIN + formChromeHeight

End Sub

Private Sub cmdExecute_Click()
    Dim intervalType As String
    Dim includeTotal As Boolean
    Dim StartYear As Long
    Dim outputRange As Range
    Dim timeSeriesData As Variant
    Dim seriesLength As Long
    Dim ws As Worksheet
    
    intervalType = Trim(txtIntervalType.Value)
    If intervalType = "" Then
        MsgBox "Please enter the Interval Type Combo (e.g., MMQHY).", vbExclamation, "Input Error"
        txtIntervalType.SetFocus
        Exit Sub
    End If
    
    includeTotal = ckbIncludeAnnualTotal.Value
    
    If Not IsNumeric(txtStartYear.Value) Then
        MsgBox "Please enter a valid Start Year.", vbExclamation, "Input Error"
        txtStartYear.SetFocus
        Exit Sub
    End If
    StartYear = CLng(txtStartYear.Value)

    On Error Resume Next
    Set outputRange = Application.Range(refOutputCell.Text)
    On Error GoTo 0
    
    If outputRange Is Nothing Then
        MsgBox "Invalid output cell reference.", vbExclamation, "Reference Error"
        refOutputCell.SetFocus
        Exit Sub
    End If
    
    Set outputRange = outputRange.Cells(1, 1)
    Set ws = outputRange.Worksheet

    ' Ensure mod_funcs exists in your project
    timeSeriesData = mod_funcs.GenerateTimeSeries(intervalType, includeTotal, StartYear)
    
    If IsArray(timeSeriesData) Then
        seriesLength = UBound(timeSeriesData)
        ws.Activate
        outputRange.Resize(1, seriesLength).Value = timeSeriesData
        MsgBox "Time series generated successfully!", vbInformation, "Success"
        Unload Me
    Else
        MsgBox "Failed to generate time series.", vbCritical, "Execution Error"
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
