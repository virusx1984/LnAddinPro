Attribute VB_Name = "frm_compare_setup"
Attribute VB_Base = "0{6753D7B7-20BD-4F7C-804B-777080770CEC}{31E73844-3361-49A3-BCBF-98A78FE39215}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' UserForm: frm_compare_setup
' Purpose: Compare Setup with Index, Reference Source, and Output Cell selection.
Option Explicit

' --- 1. Event Handler Declarations ---
Public WithEvents btnValidate As MSForms.CommandButton
Attribute btnValidate.VB_VarHelpID = -1
Public WithEvents btnReset As MSForms.CommandButton
Attribute btnReset.VB_VarHelpID = -1

' Config Role Buttons
Public WithEvents btnSetIndex As MSForms.CommandButton
Attribute btnSetIndex.VB_VarHelpID = -1
Public WithEvents btnSetCompare As MSForms.CommandButton
Attribute btnSetCompare.VB_VarHelpID = -1
Public WithEvents btnSetIgnore As MSForms.CommandButton
Attribute btnSetIgnore.VB_VarHelpID = -1
Public WithEvents btnRefUseA As MSForms.CommandButton
Attribute btnRefUseA.VB_VarHelpID = -1
Public WithEvents btnRefUseB As MSForms.CommandButton
Attribute btnRefUseB.VB_VarHelpID = -1

' Action Buttons
Public WithEvents btnRun As MSForms.CommandButton
Attribute btnRun.VB_VarHelpID = -1
Public WithEvents btnClose As MSForms.CommandButton
Attribute btnClose.VB_VarHelpID = -1

' --- 2. Control References ---
Public refRange1 As Object
Public refRange2 As Object
Public txtName1 As Object
Public txtName2 As Object
Public lstColumns As MSForms.ListBox
Public frameConfig As MSForms.Frame

' [NEW] Output Control
Public refOutput As Object

' --- 3. Layout Constants ---
Const MARGIN As Long = 10
Const CTRL_H As Long = 20
Const GAP As Long = 5
Const LBL_W As Long = 80
Const INPUT_W As Long = 180

Private Sub UserForm_Initialize()
    
    Me.Caption = "Compare Setup"
    Me.Width = 480
    Me.Height = 500 ' Increased height for Output control
    
    Dim currentTop As Long: currentTop = MARGIN
    
    ' ============================
    ' SECTION 1: RANGE SELECTION
    ' ============================
    
    ' --- Range 1 ---
    With Me.Controls.Add("Forms.Label.1", "lblRng1")
        .Caption = "1. Range A:"
        .Left = MARGIN: .Top = currentTop + 3: .Width = LBL_W
    End With
    Set refRange1 = Me.Controls.Add("RefEdit.Ctrl", "refRange1")
    With refRange1
        .Left = MARGIN + LBL_W: .Top = currentTop: .Width = INPUT_W: .Height = CTRL_H
        If TypeName(Selection) = "Range" Then .Text = Selection.Address(External:=False)
    End With
    With Me.Controls.Add("Forms.Label.1", "lblName1")
        .Caption = "Name:": .Left = MARGIN + LBL_W + INPUT_W + 10: .Top = currentTop + 3: .Width = 35
    End With
    Set txtName1 = Me.Controls.Add("Forms.TextBox.1", "txtName1")
    With txtName1
        .Left = MARGIN + LBL_W + INPUT_W + 50: .Top = currentTop: .Width = 80: .Height = CTRL_H: .Text = "BaseData"
    End With
    currentTop = currentTop + CTRL_H + GAP
    
    ' --- Range 2 ---
    With Me.Controls.Add("Forms.Label.1", "lblRng2")
        .Caption = "2. Range B:"
        .Left = MARGIN: .Top = currentTop + 3: .Width = LBL_W
    End With
    Set refRange2 = Me.Controls.Add("RefEdit.Ctrl", "refRange2")
    With refRange2
        .Left = MARGIN + LBL_W: .Top = currentTop: .Width = INPUT_W: .Height = CTRL_H
    End With
    With Me.Controls.Add("Forms.Label.1", "lblName2")
        .Caption = "Name:": .Left = MARGIN + LBL_W + INPUT_W + 10: .Top = currentTop + 3: .Width = 35
    End With
    Set txtName2 = Me.Controls.Add("Forms.TextBox.1", "txtName2")
    With txtName2
        .Left = MARGIN + LBL_W + INPUT_W + 50: .Top = currentTop: .Width = 80: .Height = CTRL_H: .Text = "TargetData"
    End With
    currentTop = currentTop + CTRL_H + GAP + 10
    
    ' --- Validate & Reset ---
    Set btnValidate = Me.Controls.Add("Forms.CommandButton.1", "btnValidate")
    With btnValidate
        .Caption = "Validate Headers": .Left = MARGIN: .Top = currentTop: .Width = 120: .Height = 24: .BackColor = &H80FF80
    End With
    Set btnReset = Me.Controls.Add("Forms.CommandButton.1", "btnReset")
    With btnReset
        .Caption = "Reset": .Left = MARGIN + 130: .Top = currentTop: .Width = 80: .Height = 24: .Enabled = False
    End With
    currentTop = currentTop + 35
    
    ' ============================
    ' SECTION 2: COLUMN CONFIG
    ' ============================
    
    Set frameConfig = Me.Controls.Add("Forms.Frame.1", "frameConfig")
    With frameConfig
        .Caption = "3. Column Config"
        .Left = MARGIN
        .Top = currentTop
        .Width = Me.InsideWidth - (MARGIN * 2)
        .Height = 230
        .Enabled = False
    End With
    
    ' ListBox
    Set lstColumns = frameConfig.Controls.Add("Forms.ListBox.1", "lstColumns")
    With lstColumns
        .Left = 10: .Top = 20: .Width = 260: .Height = 200
        .ColumnCount = 2: .ColumnWidths = "160;80"
        .MultiSelect = fmMultiSelectExtended
    End With
    
    ' --- Configuration Buttons ---
    Dim btnLeft As Long: btnLeft = 280
    Dim btnTop As Long: btnTop = 20
    
    ' Label: Key
    With frameConfig.Controls.Add("Forms.Label.1", "lblKey")
        .Caption = "Match By:": .Left = btnLeft: .Top = btnTop: .Width = 80: .Height = 15
    End With
    btnTop = btnTop + 15
    Set btnSetIndex = frameConfig.Controls.Add("Forms.CommandButton.1", "btnSetIndex")
    With btnSetIndex
        .Caption = "INDEX (Key)": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
    End With
    btnTop = btnTop + 30
    
    ' Label: Ref
    With frameConfig.Controls.Add("Forms.Label.1", "lblRef")
        .Caption = "Reference Source:": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 15
    End With
    btnTop = btnTop + 15
    Set btnRefUseA = frameConfig.Controls.Add("Forms.CommandButton.1", "btnRefUseA")
    With btnRefUseA
        .Caption = "REF (Use A)": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
        .ControlTipText = "Value from Range A"
    End With
    btnTop = btnTop + 25
    Set btnRefUseB = frameConfig.Controls.Add("Forms.CommandButton.1", "btnRefUseB")
    With btnRefUseB
        .Caption = "REF (Use B)": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
        .ControlTipText = "Value from Range B"
    End With
    btnTop = btnTop + 30
    
    ' Label: Other
    With frameConfig.Controls.Add("Forms.Label.1", "lblOth")
        .Caption = "Comparison:": .Left = btnLeft: .Top = btnTop: .Width = 80: .Height = 15
    End With
    btnTop = btnTop + 15
    Set btnSetCompare = frameConfig.Controls.Add("Forms.CommandButton.1", "btnSetCompare")
    With btnSetCompare
        .Caption = "COMPARE": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
    End With
    btnTop = btnTop + 25
    Set btnSetIgnore = frameConfig.Controls.Add("Forms.CommandButton.1", "btnSetIgnore")
    With btnSetIgnore
        .Caption = "IGNORE": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
    End With
    
    currentTop = currentTop + 240
    
    ' ============================
    ' SECTION 3: OUTPUT SELECTION [NEW]
    ' ============================
    
    With Me.Controls.Add("Forms.Label.1", "lblOut")
        .Caption = "4. Output Cell:"
        .Left = MARGIN
        .Top = currentTop + 3
        .Width = LBL_W
        .Font.Bold = True
    End With
    
    Set refOutput = Me.Controls.Add("RefEdit.Ctrl", "refOutput")
    With refOutput
        .Left = MARGIN + LBL_W
        .Top = currentTop
        .Width = INPUT_W
        .Height = CTRL_H
    End With
    
    currentTop = currentTop + 35
    
    ' ============================
    ' SECTION 4: ACTION
    ' ============================
    
    Set btnRun = Me.Controls.Add("Forms.CommandButton.1", "btnRun")
    With btnRun
        .Caption = "Run Comparison"
        .Left = Me.InsideWidth - 220
        .Top = currentTop
        .Width = 120
        .Height = 30
        .Font.Bold = True
        .Enabled = False
    End With
    
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close"
        .Left = Me.InsideWidth - 90
        .Top = currentTop
        .Width = 80
        .Height = 30
    End With
    
    Me.Height = currentTop + 70
End Sub

' --- EVENT HANDLERS ---

Private Sub btnValidate_Click()
    Dim rng1 As Range, rng2 As Range
    Dim headers1 As Variant, headers2 As Variant
    Dim i As Long
    
    On Error Resume Next
    Set rng1 = Range(refRange1.Text)
    Set rng2 = Range(refRange2.Text)
    On Error GoTo 0
    
    If rng1 Is Nothing Or rng2 Is Nothing Then MsgBox "Invalid input ranges.", vbCritical: Exit Sub
    If rng1.Columns.Count <> rng2.Columns.Count Then MsgBox "Column count mismatch.", vbCritical: Exit Sub
    
    headers1 = rng1.Rows(1).Value
    headers2 = rng2.Rows(1).Value
    
    For i = 1 To UBound(headers1, 2)
        If CStr(headers1(1, i)) <> CStr(headers2(1, i)) Then MsgBox "Header mismatch at col " & i, vbCritical: Exit Sub
    Next i
    
    ToggleInputs False
    btnReset.Enabled = True
    frameConfig.Enabled = True
    btnRun.Enabled = True
    
    lstColumns.Clear
    For i = 1 To UBound(headers1, 2)
        lstColumns.AddItem headers1(1, i)
        lstColumns.List(i - 1, 1) = "COMPARE"
    Next i
End Sub

Private Sub btnReset_Click()
    ToggleInputs True
    btnReset.Enabled = False
    frameConfig.Enabled = False
    lstColumns.Clear
    btnRun.Enabled = False
End Sub

' --- ROLE HANDLERS ---
Private Sub btnSetIndex_Click()
    UpdateColumnStatus "INDEX"
End Sub
Private Sub btnRefUseA_Click()
    UpdateColumnStatus "REF: Range A"
End Sub
Private Sub btnRefUseB_Click()
    UpdateColumnStatus "REF: Range B"
End Sub
Private Sub btnSetCompare_Click()
    UpdateColumnStatus "COMPARE"
End Sub
Private Sub btnSetIgnore_Click()
    UpdateColumnStatus "IGNORE"
End Sub

' --- RUN HANDLER ---
' --- RUN HANDLER (UPDATED) ---
Private Sub btnRun_Click()
    Dim rngA As Range
    Dim rngB As Range
    Dim outputRng As Range
    
    ' Parameter Strings
    Dim strIndex As String
    Dim strIgnore As String
    Dim strRef As String
    
    ' Function Arguments
    Dim arrIndex As Variant
    Dim arrIgnore As Variant
    Dim arrRef As Variant
    Dim bVlookupOrder As Boolean
    
    Dim i As Long
    Dim status As String
    Dim colName As String
    Dim hasRangeBPrio As Boolean
    
    ' 1. Validate Output Range
    On Error Resume Next
    Set outputRng = Range(refOutput.Text)
    Set rngA = Range(refRange1.Text)
    Set rngB = Range(refRange2.Text)
    On Error GoTo 0
    
    If outputRng Is Nothing Then
        MsgBox "Please select a valid Output Cell.", vbExclamation
        refOutput.SetFocus
        Exit Sub
    End If
    
    ' Ensure we only use the top-left cell for output
    Set outputRng = outputRng.Cells(1, 1)
    
    ' 2. Loop through ListBox to classify columns
    hasRangeBPrio = False ' Default to False (Range A priority)
    
    For i = 0 To lstColumns.listCount - 1
        colName = lstColumns.List(i, 0)
        status = lstColumns.List(i, 1)
        
        Select Case status
            Case "INDEX"
                strIndex = strIndex & colName & ","
                
            Case "IGNORE"
                strIgnore = strIgnore & colName & ","
                
            Case "REF: Range A"
                strRef = strRef & colName & ","
                ' Range A priority is default, so bVlookupOrder remains False (or unchanged)
                
            Case "REF: Range B"
                strRef = strRef & colName & ","
                ' User explicitly asked for Range B value.
                ' Note: Your function CompareExcelRanges takes a single Boolean for ALL ref columns.
                ' If ANY column requires Range B, we set the global flag to True.
                hasRangeBPrio = True
                
            ' Case "COMPARE" is implicit (columns not in Index/Ignore/Ref are compared)
        End Select
    Next i
    
    ' 3. Convert Strings to Arrays (Required by your function)
    ' We use a helper function to ensure empty strings become Empty Arrays, not Array("") containing an empty string.
    arrIndex = StringToArray(strIndex)
    arrIgnore = StringToArray(strIgnore)
    arrRef = StringToArray(strRef)
    
    ' Set the boolean order
    bVlookupOrder = hasRangeBPrio
    
    ' Check mandatory Index
    If IsEmpty(arrIndex) Then
        MsgBox "Please select at least one INDEX column.", vbExclamation
        Exit Sub
    End If
    
    ' 4. Call the Function
    Dim resultData As Variant
    
    ' Assuming CompareExcelRanges is in a Standard Module (e.g., mod_funcs)
    ' If it's in the same form code (not recommended), call it directly.
    ' Ideally, call it from the module:
    resultData = mod_funcs.CompareExcelRanges(rngA, rngB, arrIndex, arrIgnore, arrRef, bVlookupOrder)
    
    ' 5. Output the Result
    If IsArray(resultData) Then
        Dim rCount As Long, cCount As Long
        rCount = UBound(resultData, 1)
        cCount = UBound(resultData, 2)
        
        ' Check if function returned an Error Array (1D inner array check)
        ' Your function returns Array(Array("Error...")) on failure.
        ' A standard check: if row 1, col 1 starts with "Error"
        If InStr(1, CStr(resultData(1, 1)), "Error", vbTextCompare) > 0 Then
            MsgBox "Comparison Function Error: " & vbCrLf & resultData(1, 1), vbCritical
            Exit Sub
        End If
        
        ' Write to Excel
        Dim wsOut As Worksheet
        Set wsOut = outputRng.Worksheet
        wsOut.Activate
        
        ' Clear previous data area (Optional, be careful)
        ' outputRng.CurrentRegion.ClearContents
        
        ' Dump Array
        outputRng.Resize(rCount, cCount).Value = resultData
        
        ' Styling (Optional)
        outputRng.Resize(1, cCount).Font.Bold = True ' Header
        outputRng.Resize(rCount, cCount).EntireColumn.AutoFit
        
        MsgBox "Comparison Complete! Results generated at " & outputRng.Address(External:=False), vbInformation
        
        Unload Me
    Else
        MsgBox "The function did not return a valid array.", vbCritical
    End If
    
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' --- HELPERS ---
Private Sub UpdateColumnStatus(newStatus As String)
    Dim i As Long
    For i = 0 To lstColumns.listCount - 1
        If lstColumns.Selected(i) Then
            lstColumns.List(i, 1) = newStatus
            lstColumns.Selected(i) = False
        End If
    Next i
End Sub

Private Sub ToggleInputs(st As Boolean)
    refRange1.Enabled = st: refRange2.Enabled = st: txtName1.Enabled = st: txtName2.Enabled = st: btnValidate.Enabled = st
End Sub

Private Function StripComma(s As String) As String
    If Len(s) > 0 Then If Right(s, 1) = "," Then StripComma = Left(s, Len(s) - 1) Else StripComma = s
End Function

' --- Helper to convert Comma String to Array ---
Private Function StringToArray(ByVal strList As String) As Variant
    ' Removes trailing comma and returns an Array of strings.
    ' Returns Empty if string is blank.
    
    If Len(strList) = 0 Then
        StringToArray = Array() ' Return empty array
        Exit Function
    End If
    
    ' Remove trailing comma
    If Right(strList, 1) = "," Then
        strList = Left(strList, Len(strList) - 1)
    End If
    
    ' Split into array
    StringToArray = Split(strList, ",")
End Function

