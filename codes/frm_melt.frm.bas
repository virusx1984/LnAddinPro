Attribute VB_Name = "frm_melt"
Attribute VB_Base = "0{7BFC61FC-32D0-4A39-9FF9-84E9540F4FC9}{BC25BEAA-EC12-47B5-8BF6-9EA68F655083}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' UserForm: frm_melt
Option Explicit

Public WithEvents btnRun As MSForms.CommandButton
Attribute btnRun.VB_VarHelpID = -1
Public WithEvents btnCancel As MSForms.CommandButton
Attribute btnCancel.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    ' Purpose: Dynamically create and configure controls on the UserForm.

    ' --- Variable Declarations ---
    Dim yPos As Single
    Dim lblTable As MSForms.Label
    Dim refTableRange As RefEdit.RefEdit
    Dim lblIdColumns As MSForms.Label
    Dim refIdColumns As RefEdit.RefEdit
    Dim lblVarName As MSForms.Label
    Dim txtVarName As MSForms.TextBox
    Dim lblOutput As MSForms.Label
    Dim refOutputRange As RefEdit.RefEdit

    ' --- Layout Constants ---
    Const CONTROL_WIDTH As Single = 200
    Const LABEL_WIDTH As Single = 150
    Const LABEL_HEIGHT As Single = 10
    Const CONTROL_HEIGHT As Single = 20
    Const SPACING As Single = 5

    yPos = 10 ' Initial vertical position

    ' --- Section 1: Table Range ---
    Set lblTable = Me.Controls.Add("Forms.Label.1", "lblTable")
    With lblTable
        .Caption = "Table Range"
        .Left = 10
        .Top = yPos
        .Width = LABEL_WIDTH
        .Height = LABEL_HEIGHT
    End With
    yPos = yPos + LABEL_HEIGHT + SPACING

    Set refTableRange = Me.Controls.Add("RefEdit.Ctrl", "refTableRange")
    With refTableRange
        .Left = 10
        .Top = yPos
        .Width = CONTROL_WIDTH
        .Height = CONTROL_HEIGHT
        
        ' UPDATED: Set default text to current selection address if it is a Range
        If TypeName(Selection) = "Range" Then
            .Text = Selection.Address(External:=False)
        Else
            .Text = ""
        End If
    End With
    yPos = yPos + CONTROL_HEIGHT + SPACING

    ' --- Section 2: ID Columns ---
    Set lblIdColumns = Me.Controls.Add("Forms.Label.1", "lblIdColumns")
    With lblIdColumns
        .Caption = "ID Columns Range"
        .Left = 10
        .Top = yPos
        .Width = LABEL_WIDTH
        .Height = LABEL_HEIGHT
    End With
    yPos = yPos + LABEL_HEIGHT + SPACING

    Set refIdColumns = Me.Controls.Add("RefEdit.Ctrl", "refIdColumns")
    With refIdColumns
        .Left = 10
        .Top = yPos
        .Width = CONTROL_WIDTH
        .Height = CONTROL_HEIGHT
        .Text = ""
    End With
    yPos = yPos + CONTROL_HEIGHT + SPACING

    ' --- Section 3: Variable Name ---
    Set lblVarName = Me.Controls.Add("Forms.Label.1", "lblVarName")
    With lblVarName
        .Caption = "Variable Column Name"
        .Left = 10
        .Top = yPos
        .Width = LABEL_WIDTH
        .Height = LABEL_HEIGHT
    End With
    yPos = yPos + LABEL_HEIGHT + SPACING

    Set txtVarName = Me.Controls.Add("Forms.TextBox.1", "txtVarName")
    With txtVarName
        .Left = 10
        .Top = yPos
        .Width = CONTROL_WIDTH
        .Height = CONTROL_HEIGHT
        .Text = "Variable"
    End With
    yPos = yPos + CONTROL_HEIGHT + SPACING

    ' --- Section 4: Output Range ---
    Set lblOutput = Me.Controls.Add("Forms.Label.1", "lblOutput")
    With lblOutput
        .Caption = "Output Range"
        .Left = 10
        .Top = yPos
        .Width = LABEL_WIDTH
        .Height = LABEL_HEIGHT
    End With
    yPos = yPos + LABEL_HEIGHT + SPACING

    Set refOutputRange = Me.Controls.Add("RefEdit.Ctrl", "refOutputRange")
    With refOutputRange
        .Left = 10
        .Top = yPos
        .Width = CONTROL_WIDTH
        .Height = CONTROL_HEIGHT
        .Text = ""
    End With
    yPos = yPos + CONTROL_HEIGHT + SPACING + 10

    ' --- Section 5: Buttons ---
    Set btnRun = Me.Controls.Add("Forms.CommandButton.1", "btnRun")
    With btnRun
        .Caption = "Run Melt"
        .Left = 10
        .Top = yPos
        .Width = 80
        .Height = CONTROL_HEIGHT
    End With

    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With btnCancel
        .Caption = "Cancel"
        .Left = 100
        .Top = yPos
        .Width = 80
        .Height = CONTROL_HEIGHT
    End With
    yPos = yPos + CONTROL_HEIGHT + SPACING

    ' --- Final UserForm Sizing ---
    Me.Width = CONTROL_WIDTH + 20 + 10
    Me.Height = yPos + 20 + 10
    Me.Caption = "Melt Data"
End Sub

' -----------------------------------------------------------
' The btnRun_Click and btnCancel_Click subroutines remain unchanged
' -----------------------------------------------------------

Private Sub btnRun_Click()
    ' Purpose: Execute the MeltData function with user inputs
    Dim ws As Worksheet
    Dim table_range As Range, id_columns As Range, output_range As Range
    Dim var_name As String
    Dim result As Variant
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Validate inputs
    On Error Resume Next
    Set table_range = ws.Range(Me.Controls("refTableRange").Text)
    Set id_columns = ws.Range(Me.Controls("refIdColumns").Text)
    Set output_range = ws.Range(Me.Controls("refOutputRange").Text)
    var_name = Me.Controls("txtVarName").Text
    
    If table_range Is Nothing Then
        MsgBox "Invalid Table Range. Please select a valid range.", vbCritical
        Exit Sub
    End If
    If id_columns Is Nothing Then
        MsgBox "Invalid ID Columns Range. Please select a valid range.", vbCritical
        Exit Sub
    End If
    If output_range Is Nothing Then
        MsgBox "Invalid Output Range. Please select a valid cell.", vbCritical
        Exit Sub
    End If
    If Len(var_name) = 0 Then
        MsgBox "Please enter a Variable Column Name.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Ensure output_range is a single cell
    Set output_range = output_range.Cells(1, 1)
    
    ' Call the MeltData function
    result = MeltData(table_range, id_columns, var_name)
    
    ' Check if result is valid
    If IsEmpty(result) Then
        MsgBox "Error in MeltData function. Please check your inputs.", vbCritical
        Exit Sub
    End If
    
    ' Write the result to the output range
    output_range.Resize(UBound(result, 1), UBound(result, 2)).Value = result
    
    ' Inform the user and close the form
    MsgBox "Data melted successfully! Output written to " & output_range.Address, vbInformation
    Unload Me
End Sub

Private Sub btnCancel_Click()
    ' Close the form without performing any action
    Unload Me
End Sub
