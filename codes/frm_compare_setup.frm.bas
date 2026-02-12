' UserForm: frm_compare_setup
' Purpose: Complete Wizard with Range Selection, Column Roles, Ordering, Formatting, and Execution.
' UPDATED: Added Re-ordering buttons (Up/Down) and passing sorted Compare columns.
Option Explicit

' ==============================================================================
' PUBLIC CONTROL VARIABLES
' Declare these at the very top of the UserForm Code
' ==============================================================================
Public WithEvents btnValidate As MSForms.CommandButton
Public WithEvents btnReset As MSForms.CommandButton
Public WithEvents chkNewSheet As MSForms.CheckBox

' Config Role Buttons
Public WithEvents btnSetIndex As MSForms.CommandButton
Public WithEvents btnSetCompare As MSForms.CommandButton
Public WithEvents btnSetIgnore As MSForms.CommandButton
Public WithEvents btnRefUseA As MSForms.CommandButton
Public WithEvents btnRefUseB As MSForms.CommandButton
Public WithEvents btnSetFormat As MSForms.CommandButton

' Reorder Buttons
Public WithEvents btnUp As MSForms.CommandButton
Public WithEvents btnDown As MSForms.CommandButton

' Action Buttons
Public WithEvents btnRun As MSForms.CommandButton
Public WithEvents btnClose As MSForms.CommandButton

' Control References
Public refRange1 As Object
Public refRange2 As Object
Public txtName1 As Object
Public txtName2 As Object
Public refOutput As Object
Public lstColumns As MSForms.ListBox
Public frameConfig As MSForms.Frame
Public txtNewSheetName As MSForms.TextBox

' [NEW] Checkboxes
Public chkUseSheetNames As MSForms.CheckBox
Public chkFlatHeader As MSForms.CheckBox

' Layout Constants
Const MARGIN As Long = 10
Const CTRL_H As Long = 20
Const GAP As Long = 5
Const LBL_W As Long = 80
Const INPUT_W As Long = 180

' --- EVENT: Toggle Output Options ---
Private Sub chkNewSheet_Click()
    ' [FIX] Error 91 Prevention:
    ' During UserForm_Initialize, this event might trigger when setting .Value = True.
    ' At that specific moment, txtNewSheetName might not be created yet.
    ' We must check if the object exists before accessing its properties.
    If txtNewSheetName Is Nothing Or refOutput Is Nothing Then Exit Sub

    Dim isNewSheet As Boolean
    isNewSheet = chkNewSheet.Value
    
    ' 1. Enable/Disable Name TextBox
    txtNewSheetName.Enabled = isNewSheet
    txtNewSheetName.BackColor = IIf(isNewSheet, &HFFFFFF, &HCCCCCC) ' White if True, Grey if False
    
    ' 2. Enable/Disable RefEdit
    refOutput.Enabled = Not isNewSheet
End Sub

' ==============================================================================
' SUB: UserForm_Initialize
' Purpose: Builds the UI dynamically.
' ==============================================================================
Private Sub UserForm_Initialize()
    
    Me.Caption = "Compare Setup (Complete)"
    Me.Width = 550
    Me.Height = 630 ' Increased height for new output options
    
    Dim currentTop As Long: currentTop = MARGIN
    
    ' ============================
    ' PRE-CHECK SELECTION (Get Addresses & Sheet)
    ' ============================
    Dim addr1 As String
    Dim addr2 As String
    Dim selSheetName As String
    
    If TypeName(Selection) = "Range" Then
        selSheetName = "'" & Selection.Parent.Name & "'!"
        If Selection.Areas.count >= 1 Then
            addr1 = selSheetName & Selection.Areas(1).Address(External:=False)
        End If
        If Selection.Areas.count >= 2 Then
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
        .Left = MARGIN + LBL_W + INPUT_W + 50: .Top = currentTop: .Width = 80: .Height = CTRL_H
        .Text = "T1" ' Default Name
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
        .Left = MARGIN + LBL_W + INPUT_W + 50: .Top = currentTop: .Width = 80: .Height = CTRL_H
        .Text = "T2" ' Default Name
    End With
    currentTop = currentTop + CTRL_H + GAP
    
    ' --- Auto-Use Sheet Names Option ---
    Set chkUseSheetNames = Me.Controls.Add("Forms.CheckBox.1", "chkUseSheetNames")
    With chkUseSheetNames
        .Caption = "Auto-use Sheet Name if ranges are on diff sheets"
        .Left = MARGIN + LBL_W
        .Top = currentTop
        .Width = 250
        .Height = 18
        .Font.Size = 9
        .Value = True ' Default Checked
    End With
    currentTop = currentTop + 25
    
    ' --- Validate Buttons ---
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
    
    ' --- HEADERS for ListBox ---
    Dim colW1 As Double: colW1 = 150
    Dim colW2 As Double: colW2 = 90
    Dim colW3 As Double: colW3 = 80
    
    With frameConfig.Controls.Add("Forms.Label.1", "lblHdr1")
        .Caption = "Column Name": .Left = 10: .Top = 20: .Width = colW1
        .Font.bold = True: .Font.Size = 9: .ForeColor = &H8000000D
    End With
    With frameConfig.Controls.Add("Forms.Label.1", "lblHdr2")
        .Caption = "Role / Status": .Left = 10 + colW1: .Top = 20: .Width = colW2
        .Font.bold = True: .Font.Size = 9: .ForeColor = &H8000000D
    End With
    With frameConfig.Controls.Add("Forms.Label.1", "lblHdr3")
        .Caption = "Format": .Left = 10 + colW1 + colW2: .Top = 20: .Width = colW3
        .Font.bold = True: .Font.Size = 9: .ForeColor = &H8000000D
    End With
    
    ' --- LISTBOX ---
    Set lstColumns = frameConfig.Controls.Add("Forms.ListBox.1", "lstColumns")
    With lstColumns
        .Left = 10: .Top = 35
        .Width = 340: .Height = 235
        .ColumnCount = 3: .ColumnWidths = CStr(colW1) & ";" & CStr(colW2) & ";" & CStr(colW3)
        .MultiSelect = fmMultiSelectExtended
    End With
    
    ' --- BUTTONS (Right Side of Frame) ---
    Dim btnLeft As Long: btnLeft = 360
    Dim btnTop As Long: btnTop = 20
    
    ' 1. Reorder Group
    With frameConfig.Controls.Add("Forms.Label.1", "lblOrd")
        .Caption = "Order:": .Left = btnLeft: .Top = btnTop: .Width = 80: .Height = 15
    End With
    btnTop = btnTop + 15
    Set btnUp = frameConfig.Controls.Add("Forms.CommandButton.1", "btnUp")
    With btnUp
        .Caption = "Move Up": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
    End With
    btnTop = btnTop + 25
    Set btnDown = frameConfig.Controls.Add("Forms.CommandButton.1", "btnDown")
    With btnDown
        .Caption = "Move Down": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
    End With
    btnTop = btnTop + 35 ' Extra Gap
    
    ' 2. Key Group
    With frameConfig.Controls.Add("Forms.Label.1", "lblKey")
        .Caption = "Match By:": .Left = btnLeft: .Top = btnTop: .Width = 80: .Height = 15
    End With
    btnTop = btnTop + 15
    Set btnSetIndex = frameConfig.Controls.Add("Forms.CommandButton.1", "btnSetIndex")
    With btnSetIndex
        .Caption = "INDEX (Key)"
        .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
        .BackColor = &H80C0FF ' [COLOR] Orange/Gold
    End With
    btnTop = btnTop + 30
    
    ' 3. Reference Group
    With frameConfig.Controls.Add("Forms.Label.1", "lblRef")
        .Caption = "Reference Src:": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 15
    End With
    btnTop = btnTop + 15
    Set btnRefUseA = frameConfig.Controls.Add("Forms.CommandButton.1", "btnRefUseA")
    With btnRefUseA
        .Caption = "REF (Use A)"
        .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
        .BackColor = &HFFFFC0 ' [COLOR] Light Cyan
    End With
    btnTop = btnTop + 25
    Set btnRefUseB = frameConfig.Controls.Add("Forms.CommandButton.1", "btnRefUseB")
    With btnRefUseB
        .Caption = "REF (Use B)"
        .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
        .BackColor = &HFFFFC0 ' [COLOR] Light Cyan
    End With
    btnTop = btnTop + 30
    
    ' 4. Action Group
    With frameConfig.Controls.Add("Forms.Label.1", "lblOth")
        .Caption = "Action:": .Left = btnLeft: .Top = btnTop: .Width = 80: .Height = 15
    End With
    btnTop = btnTop + 15
    Set btnSetCompare = frameConfig.Controls.Add("Forms.CommandButton.1", "btnSetCompare")
    With btnSetCompare
        .Caption = "COMPARE"
        .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
        .BackColor = &H80FF80 ' [COLOR] Light Green
    End With
    btnTop = btnTop + 25
    Set btnSetIgnore = frameConfig.Controls.Add("Forms.CommandButton.1", "btnSetIgnore")
    With btnSetIgnore
        .Caption = "IGNORE"
        .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
        .BackColor = &HE0E0E0 ' [COLOR] Light Grey
    End With
    btnTop = btnTop + 30
    
    ' 5. Format Group
    Set btnSetFormat = frameConfig.Controls.Add("Forms.CommandButton.1", "btnSetFormat")
    With btnSetFormat
        .Caption = "Set Format ($%)": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
    End With
    
    currentTop = currentTop + 290
    
    ' ============================
    ' SECTION 3: OUTPUT (UPDATED)
    ' ============================
    With Me.Controls.Add("Forms.Label.1", "lblOut")
        .Caption = "4. Output Destination:": .Left = MARGIN: .Top = currentTop + 3: .Width = 150: .Font.bold = True
    End With
    currentTop = currentTop + 20
    
    ' Output Option 1: New Sheet (Default Checked)
    Set chkNewSheet = Me.Controls.Add("Forms.CheckBox.1", "chkNewSheet")
    With chkNewSheet
        .Caption = "Output to New Sheet"
        .Left = MARGIN + 10
        .Top = currentTop
        .Width = 130
        .Height = 18
        .Value = True ' Default Checked
    End With
    
    ' Sheet Name TextBox
    Set txtNewSheetName = Me.Controls.Add("Forms.TextBox.1", "txtNewSheetName")
    With txtNewSheetName
        .Left = MARGIN + 150
        .Top = currentTop
        .Width = 150
        .Height = 18
        .Text = "CmpResult"
        .Enabled = True
    End With
    currentTop = currentTop + 25
    
    ' Output Option 2: Existing Cell (Disabled by default)
    With Me.Controls.Add("Forms.Label.1", "lblRefOut")
        .Caption = "Or select cell:": .Left = MARGIN + 10: .Top = currentTop + 3: .Width = 80
    End With
    Set refOutput = Me.Controls.Add("RefEdit.Ctrl", "refOutput")
    With refOutput
        .Left = MARGIN + 90
        .Top = currentTop
        .Width = 210
        .Height = CTRL_H
        .Enabled = False ' Disabled because "New Sheet" is checked
    End With
    currentTop = currentTop + 30
    
    ' ============================
    ' OPTION: FLAT HEADER
    ' ============================
    Set chkFlatHeader = Me.Controls.Add("Forms.CheckBox.1", "chkFlatHeader")
    With chkFlatHeader
        .Caption = "Flat Header (e.g., T1_Col1)"
        .Left = MARGIN + 10
        .Top = currentTop
        .Width = 200
        .Height = 20
        .Value = True ' Default Checked
    End With
    currentTop = currentTop + 35
    
    ' ============================
    ' SECTION 4: ACTION
    ' ============================
    Set btnRun = Me.Controls.Add("Forms.CommandButton.1", "btnRun")
    With btnRun
        .Caption = "Run Comparison": .Left = Me.InsideWidth - 220: .Top = currentTop: .Width = 120: .Height = 30: .Font.bold = True: .Enabled = False
    End With
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close": .Left = Me.InsideWidth - 90: .Top = currentTop: .Width = 80: .Height = 30
    End With
    
    ' --- Sync Initial State ---
    ' Ensure UI reflects the default True value of NewSheet
    txtNewSheetName.Enabled = True
    txtNewSheetName.BackColor = &HFFFFFF
    refOutput.Enabled = False
    
    Me.Height = currentTop + 80
End Sub

' --- EVENT HANDLERS ---
Private Sub btnValidate_Click()
    Dim rng1 As Range, rng2 As Range
    Dim headers1 As Variant, headers2 As Variant
    Dim i As Long
    Dim dataVal As Variant
    Dim isNum As Boolean
    
    ' Error handling
    On Error Resume Next
    Set rng1 = Range(refRange1.Text)
    Set rng2 = Range(refRange2.Text)
    On Error GoTo 0
    
    ' --- Basic Validation ---
    If rng1 Is Nothing Or rng2 Is Nothing Then MsgBox "Invalid ranges.", vbCritical: Exit Sub
    If rng1.Columns.count <> rng2.Columns.count Then MsgBox "Column count mismatch.", vbCritical: Exit Sub
    
    ' ==============================================================================
    ' [UPDATED] Name Handling Logic
    ' 1. If "Auto-use Sheet Names" is CHECKED and sheets are different -> Overwrite.
    ' 2. If UNCHECKED -> Keep existing text.
    ' 3. Empty Check -> Ensure names are not empty.
    ' ==============================================================================
    
    ' Check logic: If checked and sheets differ, force update
    If chkUseSheetNames.Value = True Then
        If rng1.Parent.Name <> rng2.Parent.Name Then
            txtName1.Text = rng1.Parent.Name
            txtName2.Text = rng2.Parent.Name
        End If
    End If
    
    ' Validation: Names cannot be empty
    If Trim(txtName1.Text) = "" Then
        If rng1.Parent.Name <> rng2.Parent.Name Then
             txtName1.Text = rng1.Parent.Name
        Else
             txtName1.Text = "T1"
        End If
    End If
    
    If Trim(txtName2.Text) = "" Then
        If rng1.Parent.Name <> rng2.Parent.Name Then
             txtName2.Text = rng2.Parent.Name
        Else
             txtName2.Text = "T2"
        End If
    End If
    ' ==============================================================================

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
        
        If rng1.Rows.count > 1 Then
            dataVal = rng1.Cells(2, i).Value
            isNum = IsNumeric(dataVal) And Not IsEmpty(dataVal)
        Else
            isNum = False
        End If
        
        If isNum Then
            lstColumns.List(i - 1, 1) = "COMPARE"
            lstColumns.List(i - 1, 2) = "#,##0.00_-;[Red]-#,##0.00_-;""-""_-;@"
        Else
            lstColumns.List(i - 1, 1) = "INDEX"
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

' --- [NEW] REORDER BUTTONS HANDLERS ---
Private Sub btnUp_Click()
    MoveListBoxItem -1
End Sub

Private Sub btnDown_Click()
    MoveListBoxItem 1
End Sub

Private Sub MoveListBoxItem(offset As Long)
    Dim i As Long, j As Long, k As Long
    Dim tempVal As String
    
    With lstColumns
        For i = 0 To .listCount - 1
            If .Selected(i) Then
                ' Check bounds
                If (offset < 0 And i > 0) Or (offset > 0 And i < .listCount - 1) Then
                    j = i + offset
                    ' Swap columns
                    For k = 0 To .ColumnCount - 1
                        tempVal = .List(i, k)
                        .List(i, k) = .List(j, k)
                        .List(j, k) = tempVal
                    Next k
                    
                    ' Update Selection
                    .Selected(i) = False
                    .Selected(j) = True
                    Exit For ' Move one item at a time (or remove to support multi-move logic)
                End If
            End If
        Next i
    End With
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
    Dim strIndex As String, strIgnore As String, strRef As String, strCompare As String
    Dim arrIndex As Variant, arrIgnore As Variant, arrRef As Variant, arrCompare As Variant
    Dim dictRefDirs As Object: Set dictRefDirs = CreateObject("Scripting.Dictionary")
    Dim finalRefDirs As Variant
    Dim i As Long, status As String, colName As String, colFmt As String
    Dim dictFormats As Object: Set dictFormats = CreateObject("Scripting.Dictionary")
    
    ' Get Table Names
    Dim name1 As String, name2 As String
    name1 = txtName1.Text: name2 = txtName2.Text
    If name1 = "" Then name1 = "T1"
    If name2 = "" Then name2 = "T2"
    
    ' Set Ranges
    On Error Resume Next
    Set rngA = Range(refRange1.Text)
    Set rngB = Range(refRange2.Text)
    On Error GoTo 0
    
    If rngA Is Nothing Or rngB Is Nothing Then
        MsgBox "Invalid Ranges selected.", vbCritical
        Exit Sub
    End If
    
    ' ============================================================
    ' [FIXED] OUTPUT LOGIC: NEW SHEET CREATION
    ' ============================================================
    If chkNewSheet.Value = True Then
        ' --- OPTION A: NEW SHEET ---
        Dim baseName As String
        Dim finalName As String
        Dim counter As Integer
        Dim targetWB As Workbook ' Define the target workbook
        Dim wsNew As Worksheet
        Dim wsCheck As Worksheet
        
        ' 1. Determine Target Workbook (Same as Range A)
        Set targetWB = rngA.Worksheet.Parent
        
        ' 2. Get Base Name
        baseName = Trim(txtNewSheetName.Text)
        If baseName = "" Then baseName = "CmpResult"
        
        finalName = baseName
        counter = 1
        
        ' 3. Check for Name Collision (Auto-Increment)
        Do
            Set wsCheck = Nothing
            On Error Resume Next
            Set wsCheck = targetWB.Sheets(finalName)
            On Error GoTo 0
            
            If Not wsCheck Is Nothing Then
                ' Name exists, increment count
                finalName = baseName & counter
                counter = counter + 1
            Else
                ' Unique Name found
                Exit Do
            End If
        Loop
        
        ' 4. Create Sheet in the Target Workbook
        Set wsNew = targetWB.Worksheets.Add(After:=targetWB.Worksheets(targetWB.Worksheets.count))
        
        ' 5. Rename Sheet (Handle errors in case name is invalid)
        On Error Resume Next
        wsNew.Name = finalName
        If Err.Number <> 0 Then
            MsgBox "Error renaming sheet to '" & finalName & "'. Using default name.", vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
        
        ' 6. Set Output to A1 of the new sheet
        Set outputRng = wsNew.Range("A1")
        
    Else
        ' --- OPTION B: EXISTING CELL ---
        On Error Resume Next
        Set outputRng = Range(refOutput.Text)
        On Error GoTo 0
        
        If outputRng Is Nothing Then MsgBox "Select valid output cell.", vbExclamation: Exit Sub
        Set outputRng = outputRng.Cells(1, 1)
    End If
    ' ============================================================
    
    ' --- LOOP LISTBOX (Config) ---
    For i = 0 To lstColumns.listCount - 1
        colName = lstColumns.List(i, 0)
        status = lstColumns.List(i, 1)
        colFmt = lstColumns.List(i, 2)
        If Len(colFmt) > 0 And LCase(colFmt) <> "general" Then dictFormats.item(colName) = colFmt
        Select Case status
            Case "INDEX": strIndex = strIndex & colName & ","
            Case "IGNORE": strIgnore = strIgnore & colName & ","
            Case "REF: Range A": strRef = strRef & colName & ",": dictRefDirs.item(colName) = False
            Case "REF: Range B": strRef = strRef & colName & ",": dictRefDirs.item(colName) = True
            Case "COMPARE": strCompare = strCompare & colName & ","
        End Select
    Next i
    
    arrIndex = StringToArray(strIndex)
    arrIgnore = StringToArray(strIgnore)
    arrRef = StringToArray(strRef)
    arrCompare = StringToArray(strCompare)
    If dictRefDirs.count > 0 Then Set finalRefDirs = dictRefDirs Else finalRefDirs = Empty
    
    If IsEmpty(arrIndex) Or UBound(arrIndex) = -1 Then MsgBox "Select at least one INDEX column.", vbExclamation: Exit Sub
    
    ' --- Flat Header Check ---
    Dim isFlat As Boolean: isFlat = chkFlatHeader.Value
    
    ' --- CALL MAIN FUNCTION ---
    Dim resultData As Variant
    resultData = mod_funcs.CompareExcelRanges( _
        rngA, rngB, arrIndex, arrIgnore, arrRef, finalRefDirs, arrCompare, name1, name2, isFlat _
    )
    
    ' --- OUTPUT RESULTS ---
    If IsArray(resultData) Then
        Dim rCount As Long, cCount As Long
        rCount = UBound(resultData, 1)
        cCount = UBound(resultData, 2)
        
        If InStr(1, CStr(resultData(1, 1)), "Error", vbTextCompare) > 0 Then
            MsgBox "Error: " & resultData(1, 1), vbCritical: Exit Sub
        End If
        
        ' Activate the target sheet (New or Existing)
        outputRng.Worksheet.Activate
        outputRng.Resize(rCount, cCount).Value = resultData
        outputRng.Resize(1, cCount).Font.bold = True
        
        ' --- APPLY FORMATTING ---
        If dictFormats.count > 0 Then
            Dim headerRng As Range, cell As Range
            Dim headerRowIndex As Long, dataRowCount As Long
            
            If isFlat Then
                headerRowIndex = 0: dataRowCount = rCount - 1
            Else
                headerRowIndex = 1: dataRowCount = rCount - 2
            End If
            
            If dataRowCount > 0 Then
                Set headerRng = outputRng.offset(headerRowIndex, 0).Resize(1, cCount)
                For Each cell In headerRng.Cells
                    Dim fmtKey As String: fmtKey = cell.Value
                    If dictFormats.Exists(fmtKey) Then
                        cell.offset(1, 0).Resize(dataRowCount, 1).NumberFormat = dictFormats(fmtKey)
                    Else
                        Dim prefixes As Variant: prefixes = Array(name1 & "_", name2 & "_", "Diff_")
                        Dim p As Variant
                        For Each p In prefixes
                            If InStr(1, fmtKey, p, vbTextCompare) = 1 Then
                                fmtKey = Mid(fmtKey, Len(p) + 1)
                                Exit For
                            End If
                        Next p
                        If dictFormats.Exists(fmtKey) Then cell.offset(1, 0).Resize(dataRowCount, 1).NumberFormat = dictFormats(fmtKey)
                    End If
                Next cell
            End If
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

Private Sub UpdateColumnStatus(newStatus As String)
    Dim i As Long
    For i = 0 To lstColumns.listCount - 1
        If lstColumns.Selected(i) Then
            lstColumns.List(i, 1) = newStatus
            lstColumns.Selected(i) = False
            Select Case newStatus
                Case "INDEX": lstColumns.List(i, 2) = "@"
                Case "REF: Range A": lstColumns.List(i, 2) = "@"
                Case "REF: Range B": lstColumns.List(i, 2) = "@"
                Case "COMPARE": lstColumns.List(i, 2) = "#,##0.00_-;[Red]-#,##0.00_-;""-""_-;@"
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