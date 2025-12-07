Attribute VB_Name = "frm_compare_setup"
Attribute VB_Base = "0{0C55C441-DDEE-4A37-8482-D65B41D73148}{6F04D38A-C6B4-46D8-BE6F-42C2ADB30350}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' UserForm: frm_compare_setup
' Purpose: Complete Wizard with Range Selection, Column Roles, Ordering, Formatting, and Execution.
' UPDATED: Auto-fill Range A and B from multi-selection (Areas).
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
Public WithEvents btnSetFormat As MSForms.CommandButton
Attribute btnSetFormat.VB_VarHelpID = -1

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
Public refOutput As Object
Public lstColumns As MSForms.ListBox
Public frameConfig As MSForms.Frame

' --- 3. Layout Constants ---
Const MARGIN As Long = 10
Const CTRL_H As Long = 20
Const GAP As Long = 5
Const LBL_W As Long = 80
Const INPUT_W As Long = 180

Private Sub UserForm_Initialize()
    
    Me.Caption = "Compare Setup (Complete)"
    Me.Width = 550
    Me.Height = 550
    
    Dim currentTop As Long: currentTop = MARGIN
    
    ' ============================
    ' PRE-CHECK SELECTION (Updated with Sheet Names)
    ' ============================
    Dim addr1 As String
    Dim addr2 As String
    Dim selSheetName As String
    
    ' ¼ì²éµ±Ç°Ñ¡ÖÐÁËÊ²Ã´
    If TypeName(Selection) = "Range" Then
        ' »ñÈ¡µ±Ç°Ñ¡ÖÐÇøÓòËùÔÚµÄ¹¤×÷±íÃû³Æ (´¦Àí´ø¿Õ¸ñµÄÃû×ÖÐèÒª¼ÓÉÏµ¥ÒýºÅ)
        selSheetName = "'" & Selection.Parent.Name & "'!"
        
        If Selection.Areas.count >= 1 Then
            ' Æ´½Ó SheetÃû + µØÖ·
            addr1 = selSheetName & Selection.Areas(1).Address(External:=False)
        End If
        
        If Selection.Areas.count >= 2 Then
            ' Æ´½Ó SheetÃû + µØÖ·
            addr2 = selSheetName & Selection.Areas(2).Address(External:=False)
        End If
    End If
    
    ' ============================
    ' SECTION 1: RANGE SELECTION
    ' ============================
    
    ' --- Range A ---
    With Me.Controls.Add("Forms.Label.1", "lblRng1")
        .Caption = "1. Range A:": .Left = MARGIN: .Top = currentTop + 3: .Width = LBL_W
    End With
    Set refRange1 = Me.Controls.Add("RefEdit.Ctrl", "refRange1")
    With refRange1
        .Left = MARGIN + LBL_W: .Top = currentTop: .Width = INPUT_W: .Height = CTRL_H
        .Text = addr1
    End With
    With Me.Controls.Add("Forms.Label.1", "lblName1")
        .Caption = "Name:": .Left = MARGIN + LBL_W + INPUT_W + 10: .Top = currentTop + 3: .Width = 35
    End With
    Set txtName1 = Me.Controls.Add("Forms.TextBox.1", "txtName1")
    With txtName1
        .Left = MARGIN + LBL_W + INPUT_W + 50: .Top = currentTop: .Width = 80: .Height = CTRL_H: .Text = "BaseData"
    End With
    currentTop = currentTop + CTRL_H + GAP
    
    ' --- Range B ---
    With Me.Controls.Add("Forms.Label.1", "lblRng2")
        .Caption = "2. Range B:": .Left = MARGIN: .Top = currentTop + 3: .Width = LBL_W
    End With
    Set refRange2 = Me.Controls.Add("RefEdit.Ctrl", "refRange2")
    With refRange2
        .Left = MARGIN + LBL_W: .Top = currentTop: .Width = INPUT_W: .Height = CTRL_H
        .Text = addr2
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
        .Caption = "3. Column Config (Role, Order & Format)"
        .Left = MARGIN
        .Top = currentTop
        .Width = Me.InsideWidth - (MARGIN * 2)
        .Height = 280
        .Enabled = False
    End With
    
    ' --- HEADERS ---
    Dim colW1 As Double: colW1 = 150
    Dim colW2 As Double: colW2 = 90
    Dim colW3 As Double: colW3 = 80
    
    With frameConfig.Controls.Add("Forms.Label.1", "lblHdr1")
        .Caption = "Column Name": .Left = 10: .Top = 20: .Width = colW1
        .Font.Bold = True: .Font.Size = 9: .ForeColor = &H8000000D
    End With
    With frameConfig.Controls.Add("Forms.Label.1", "lblHdr2")
        .Caption = "Role / Status": .Left = 10 + colW1: .Top = 20: .Width = colW2
        .Font.Bold = True: .Font.Size = 9: .ForeColor = &H8000000D
    End With
    With frameConfig.Controls.Add("Forms.Label.1", "lblHdr3")
        .Caption = "Format": .Left = 10 + colW1 + colW2: .Top = 20: .Width = colW3
        .Font.Bold = True: .Font.Size = 9: .ForeColor = &H8000000D
    End With
    
    ' --- LISTBOX ---
    Set lstColumns = frameConfig.Controls.Add("Forms.ListBox.1", "lstColumns")
    With lstColumns
        .Left = 10: .Top = 35: .Width = 340: .Height = 235
        .ColumnCount = 3: .ColumnWidths = CStr(colW1) & ";" & CStr(colW2) & ";" & CStr(colW3)
        .MultiSelect = fmMultiSelectExtended
    End With
    
    ' --- Buttons (Right Side) ---
    Dim btnLeft As Long: btnLeft = 360
    Dim btnTop As Long: btnTop = 20
    
    With frameConfig.Controls.Add("Forms.Label.1", "lblKey")
        .Caption = "Match By:": .Left = btnLeft: .Top = btnTop: .Width = 80: .Height = 15
    End With
    btnTop = btnTop + 15
    Set btnSetIndex = frameConfig.Controls.Add("Forms.CommandButton.1", "btnSetIndex")
    With btnSetIndex
        .Caption = "INDEX (Key)": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
    End With
    btnTop = btnTop + 30
    
    With frameConfig.Controls.Add("Forms.Label.1", "lblRef")
        .Caption = "Reference Src:": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 15
    End With
    btnTop = btnTop + 15
    Set btnRefUseA = frameConfig.Controls.Add("Forms.CommandButton.1", "btnRefUseA")
    With btnRefUseA
        .Caption = "REF (Use A)": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
    End With
    btnTop = btnTop + 25
    Set btnRefUseB = frameConfig.Controls.Add("Forms.CommandButton.1", "btnRefUseB")
    With btnRefUseB
        .Caption = "REF (Use B)": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
    End With
    btnTop = btnTop + 30
    
    With frameConfig.Controls.Add("Forms.Label.1", "lblOth")
        .Caption = "Action:": .Left = btnLeft: .Top = btnTop: .Width = 80: .Height = 15
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
    btnTop = btnTop + 30
    
    Set btnSetFormat = frameConfig.Controls.Add("Forms.CommandButton.1", "btnSetFormat")
    With btnSetFormat
        .Caption = "Set Format ($%)": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
    End With
    
    currentTop = currentTop + 290
    
    ' ============================
    ' SECTION 3: OUTPUT
    ' ============================
    With Me.Controls.Add("Forms.Label.1", "lblOut")
        .Caption = "4. Output Cell:": .Left = MARGIN: .Top = currentTop + 3: .Width = LBL_W: .Font.Bold = True
    End With
    Set refOutput = Me.Controls.Add("RefEdit.Ctrl", "refOutput")
    With refOutput
        .Left = MARGIN + LBL_W: .Top = currentTop: .Width = INPUT_W: .Height = CTRL_H
    End With
    currentTop = currentTop + 35
    
    ' ============================
    ' SECTION 4: ACTION
    ' ============================
    Set btnRun = Me.Controls.Add("Forms.CommandButton.1", "btnRun")
    With btnRun
        .Caption = "Run Comparison": .Left = Me.InsideWidth - 220: .Top = currentTop: .Width = 120: .Height = 30: .Font.Bold = True: .Enabled = False
    End With
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close": .Left = Me.InsideWidth - 90: .Top = currentTop: .Width = 80: .Height = 30
    End With
    
    Me.Height = currentTop + 70
End Sub

' --- EVENT HANDLERS ---

Private Sub btnValidate_Click()
    Dim rng1 As Range, rng2 As Range
    Dim headers1 As Variant, headers2 As Variant
    Dim i As Long
    Dim dataVal As Variant
    Dim isNum As Boolean
    
    On Error Resume Next
    Set rng1 = Range(refRange1.Text)
    Set rng2 = Range(refRange2.Text)
    On Error GoTo 0
    
    ' --- Basic Validation ---
    If rng1 Is Nothing Or rng2 Is Nothing Then MsgBox "Invalid ranges.", vbCritical: Exit Sub
    If rng1.Columns.count <> rng2.Columns.count Then MsgBox "Column count mismatch.", vbCritical: Exit Sub
    
    ' Get Headers
    headers1 = rng1.Rows(1).Value
    headers2 = rng2.Rows(1).Value
    
    ' Verify Header Consistency
    For i = 1 To UBound(headers1, 2)
        If CStr(headers1(1, i)) <> CStr(headers2(1, i)) Then MsgBox "Header mismatch at col " & i, vbCritical: Exit Sub
    Next i
    
    ' --- Validation Passed, Enable Config Section ---
    ToggleInputs False
    btnReset.Enabled = True
    frameConfig.Enabled = True
    btnRun.Enabled = True
    
    ' --- Populate ListBox & Auto-Detect Roles ---
    lstColumns.Clear
    For i = 1 To UBound(headers1, 2)
        lstColumns.AddItem headers1(1, i)
        
        ' Logic: Check the 2nd row (data row) of Range A
        ' If rows > 1, check 2nd row; else (header only) assume non-numeric
        If rng1.Rows.count > 1 Then
            dataVal = rng1.Cells(2, i).Value
            ' Check if numeric and not empty
            isNum = IsNumeric(dataVal) And Not IsEmpty(dataVal)
        Else
            isNum = False
        End If
        
        If isNum Then
            ' === Numeric: Default to COMPARE ===
            lstColumns.List(i - 1, 1) = "COMPARE"
            ' Set standard number format (thousands separator, red for negative)
            lstColumns.List(i - 1, 2) = "#,##0.00_-;[Red]-#,##0.00_-;""-""_-;@"
        Else
            ' === Non-Numeric: Default to INDEX ===
            lstColumns.List(i - 1, 1) = "INDEX"
            ' Set text format
            lstColumns.List(i - 1, 2) = "@"
        End If
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

' --- FORMAT HANDLER ---
Private Sub btnSetFormat_Click()
    Dim strFormat As String
    Dim i As Long
    
    strFormat = InputBox("Enter Excel Number Format string:" & vbCrLf & "e.g., 0.00, $#,##0.00, 0%", "Set Column Format", "0.00")
    If StrPtr(strFormat) = 0 Then Exit Sub
    If strFormat = "" Then strFormat = "General"
    
    For i = 0 To lstColumns.listCount - 1
        If lstColumns.Selected(i) Then
            lstColumns.List(i, 2) = strFormat
            lstColumns.Selected(i) = False
        End If
    Next i
End Sub

' --- RUN HANDLER ---
Private Sub btnRun_Click()
    Dim rngA As Range, rngB As Range, outputRng As Range
    Dim strIndex As String, strIgnore As String, strRef As String
    
    ' Arrays to pass to function
    Dim arrIndex As Variant, arrIgnore As Variant, arrRef As Variant
    
    ' Dictionary for Reference Directions
    Dim dictRefDirs As Object
    Set dictRefDirs = CreateObject("Scripting.Dictionary")
    
    Dim finalRefDirs As Variant
    
    Dim i As Long, status As String, colName As String, colFmt As String
    Dim dictFormats As Object: Set dictFormats = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set outputRng = Range(refOutput.Text)
    Set rngA = Range(refRange1.Text)
    Set rngB = Range(refRange2.Text)
    On Error GoTo 0
    
    If outputRng Is Nothing Then MsgBox "Select valid output cell.", vbExclamation: Exit Sub
    Set outputRng = outputRng.Cells(1, 1)
    
    ' --- LOOP THROUGH LISTBOX TO BUILD CONFIG ---
    For i = 0 To lstColumns.listCount - 1
        colName = lstColumns.List(i, 0)
        status = lstColumns.List(i, 1)
        colFmt = lstColumns.List(i, 2)
        
        ' Store format for later
        If Len(colFmt) > 0 And LCase(colFmt) <> "general" Then dictFormats.item(colName) = colFmt
        
        Select Case status
            Case "INDEX"
                strIndex = strIndex & colName & ","
                
            Case "IGNORE"
                strIgnore = strIgnore & colName & ","
                
            Case "REF: Range A"
                strRef = strRef & colName & ","
                dictRefDirs.item(colName) = False
                
            Case "REF: Range B"
                strRef = strRef & colName & ","
                dictRefDirs.item(colName) = True
        End Select
    Next i
    
    ' Convert CSV strings to Arrays
    arrIndex = StringToArray(strIndex)
    arrIgnore = StringToArray(strIgnore)
    arrRef = StringToArray(strRef)
    
    If dictRefDirs.count > 0 Then
        Set finalRefDirs = dictRefDirs
    Else
        finalRefDirs = Empty
    End If
    
    ' Check required inputs
    If IsEmpty(arrIndex) Then MsgBox "Select at least one INDEX column.", vbExclamation: Exit Sub
    If UBound(arrIndex) = -1 Then MsgBox "Select at least one INDEX column.", vbExclamation: Exit Sub
    
    ' --- CALL MAIN FUNCTION ---
    Dim resultData As Variant
    ' Check if mod_funcs exists, assume it does based on your context
    resultData = mod_funcs.CompareExcelRanges( _
        rngA, _
        rngB, _
        arrIndex, _
        arrIgnore, _
        arrRef, _
        finalRefDirs _
    )
    
    ' --- OUTPUT RESULTS ---
    If IsArray(resultData) Then
        Dim rCount As Long, cCount As Long
        rCount = UBound(resultData, 1)
        cCount = UBound(resultData, 2)
        
        If InStr(1, CStr(resultData(1, 1)), "Error", vbTextCompare) > 0 Then
            MsgBox "Error: " & resultData(1, 1), vbCritical: Exit Sub
        End If
        
        outputRng.Worksheet.Activate
        outputRng.Resize(rCount, cCount).Value = resultData
        outputRng.Resize(1, cCount).Font.Bold = True
        
        ' APPLY FORMATTING
        If dictFormats.count > 0 And rCount > 2 Then
            Dim headerRng As Range, cell As Range
            Set headerRng = outputRng.Offset(1, 0).Resize(1, cCount)
            For Each cell In headerRng.Cells
                If dictFormats.Exists(cell.Value) Then
                    cell.Offset(1, 0).Resize(rCount - 2, 1).NumberFormat = dictFormats(cell.Value)
                End If
                
                Dim baseName As String
                baseName = cell.Value
                If InStr(baseName, "_T1") > 0 Then baseName = Replace(baseName, "_T1", "")
                If InStr(baseName, "_T2") > 0 Then baseName = Replace(baseName, "_T2", "")
                If InStr(baseName, "_Diff") > 0 Then baseName = Replace(baseName, "_Diff", "")
                
                If dictFormats.Exists(baseName) Then
                     cell.Offset(1, 0).Resize(rCount - 2, 1).NumberFormat = dictFormats(baseName)
                End If
            Next cell
        End If
        
        MsgBox "Comparison Complete!", vbInformation
        Unload Me
    Else
        MsgBox "Function failed to return array.", vbCritical
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
            Select Case newStatus
                Case "INDEX"
                    lstColumns.List(i, 2) = "@"
                Case "REF: Range A"
                    lstColumns.List(i, 2) = "@"
                Case "REF: Range B"
                    lstColumns.List(i, 2) = "@"
                Case "COMPARE"
                    lstColumns.List(i, 2) = "#,##0_-;[Color3]-#,##0_-;""-""_-;@"
                    
            End Select
                
        End If
    Next i
End Sub

Private Sub ToggleInputs(st As Boolean)
    refRange1.Enabled = st: refRange2.Enabled = st: txtName1.Enabled = st: txtName2.Enabled = st: btnValidate.Enabled = st
End Sub

Private Function StringToArray(ByVal strList As String) As Variant
    If Len(strList) = 0 Then
        StringToArray = Empty
        Exit Function
    End If
    
    If Right(strList, 1) = "," Then strList = Left(strList, Len(strList) - 1)
    StringToArray = Split(strList, ",")
End Function

