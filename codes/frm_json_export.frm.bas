Attribute VB_Name = "frm_json_export"
Attribute VB_Base = "0{8AE76E9D-3720-49FE-8FD8-97D33542B133}{C3CE3BB8-C101-47CE-82D6-E5107040F964}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' UserForm: frm_json_export
' Purpose: Dynamic Wizard with File Dialog and improved layout.
Option Explicit

' --- 1. Event Handler Declarations ---
Public WithEvents btnValidate As MSForms.CommandButton
Attribute btnValidate.VB_VarHelpID = -1
Public WithEvents btnReset As MSForms.CommandButton
Attribute btnReset.VB_VarHelpID = -1

' Config Role Buttons
Public WithEvents btnSetCate As MSForms.CommandButton
Attribute btnSetCate.VB_VarHelpID = -1
Public WithEvents btnSetAgg As MSForms.CommandButton
Attribute btnSetAgg.VB_VarHelpID = -1
Public WithEvents btnSetFormat As MSForms.CommandButton
Attribute btnSetFormat.VB_VarHelpID = -1

' File & Action Buttons
Public WithEvents btnBrowse As MSForms.CommandButton
Attribute btnBrowse.VB_VarHelpID = -1
Public WithEvents btnRun As MSForms.CommandButton
Attribute btnRun.VB_VarHelpID = -1
Public WithEvents btnClose As MSForms.CommandButton
Attribute btnClose.VB_VarHelpID = -1

' --- 2. Control References ---
Public refRangeSource As Object
Public txtOutputName As Object
Public lstColumns As MSForms.ListBox
Public frameConfig As MSForms.Frame

' --- 3. Layout Constants ---
Const MARGIN As Long = 10
Const CTRL_H As Long = 20
Const GAP As Long = 5
Const LBL_W As Long = 80
Const INPUT_W As Long = 220 ' Increased width for file path

Private Sub UserForm_Initialize()
    
    Me.Caption = "JSON Export Wizard"
    Me.Width = 550
    Me.Height = 550 ' Increased height for better spacing
    
    Dim currentTop As Long: currentTop = MARGIN
    
    ' ============================
    ' SECTION 1: RANGE SELECTION
    ' ============================
    With Me.Controls.Add("Forms.Label.1", "lblRngSrc")
        .Caption = "1. Data Range:": .Left = MARGIN: .Top = currentTop + 3: .Width = LBL_W
    End With
    
    ' Try RefEdit, fallback to TextBox
    On Error Resume Next
    Set refRangeSource = Me.Controls.Add("RefEdit.Ctrl", "refRangeSource")
    If Err.Number <> 0 Then Set refRangeSource = Me.Controls.Add("Forms.TextBox.1", "refRangeSource")
    On Error GoTo 0
    
    With refRangeSource
        .Left = MARGIN + LBL_W: .Top = currentTop: .Width = INPUT_W - 40: .Height = CTRL_H
        If TypeName(Selection) = "Range" Then .Text = Selection.Address(External:=False)
    End With
    
    currentTop = currentTop + CTRL_H + GAP + 10
    
    ' Validate / Reset
    Set btnValidate = Me.Controls.Add("Forms.CommandButton.1", "btnValidate")
    With btnValidate
        .Caption = "Load Headers": .Left = MARGIN: .Top = currentTop: .Width = 120: .Height = 24: .BackColor = &H80FF80
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
        .Caption = "2. Column Config (Type & Format)"
        .Left = MARGIN: .Top = currentTop: .Width = Me.InsideWidth - (MARGIN * 2): .Height = 280: .Enabled = False
    End With
    
    ' ListBox Headers
    Dim colW1 As Double: colW1 = 150
    Dim colW2 As Double: colW2 = 90
    Dim colW3 As Double: colW3 = 80
    
    With frameConfig.Controls.Add("Forms.Label.1", "lblHdr1")
        .Caption = "Column Name": .Left = 10: .Top = 20: .Width = colW1: .Font.Bold = True: .ForeColor = &H8000000D
    End With
    With frameConfig.Controls.Add("Forms.Label.1", "lblHdr2")
        .Caption = "Type": .Left = 10 + colW1: .Top = 20: .Width = colW2: .Font.Bold = True: .ForeColor = &H8000000D
    End With
    With frameConfig.Controls.Add("Forms.Label.1", "lblHdr3")
        .Caption = "Format": .Left = 10 + colW1 + colW2: .Top = 20: .Width = colW3: .Font.Bold = True: .ForeColor = &H8000000D
    End With
    
    Set lstColumns = frameConfig.Controls.Add("Forms.ListBox.1", "lstColumns")
    With lstColumns
        .Left = 10: .Top = 35: .Width = 340: .Height = 235
        .ColumnCount = 3: .ColumnWidths = CStr(colW1) & ";" & CStr(colW2) & ";" & CStr(colW3)
        .MultiSelect = fmMultiSelectExtended
    End With
    
    ' Config Buttons
    Dim btnLeft As Long: btnLeft = 360
    Dim btnTop As Long: btnTop = 35
    
    With frameConfig.Controls.Add("Forms.Label.1", "lblType")
        .Caption = "Set Type:": .Left = btnLeft: .Top = btnTop: .Width = 80: .Height = 15
    End With
    btnTop = btnTop + 15
    Set btnSetCate = frameConfig.Controls.Add("Forms.CommandButton.1", "btnSetCate")
    With btnSetCate
        .Caption = "CATEGORY": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 24
    End With
    btnTop = btnTop + 30
    Set btnSetAgg = frameConfig.Controls.Add("Forms.CommandButton.1", "btnSetAgg")
    With btnSetAgg
        .Caption = "AGGREGATION": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 24
    End With
    btnTop = btnTop + 40
    
    With frameConfig.Controls.Add("Forms.Label.1", "lblFmt")
        .Caption = "Formatting:": .Left = btnLeft: .Top = btnTop: .Width = 80: .Height = 15
    End With
    btnTop = btnTop + 15
    Set btnSetFormat = frameConfig.Controls.Add("Forms.CommandButton.1", "btnSetFormat")
    With btnSetFormat
        .Caption = "Set Format...": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 24
    End With
    
    currentTop = currentTop + 290
    
    ' ============================
    ' SECTION 3: FILE SELECTION
    ' ============================
    With Me.Controls.Add("Forms.Label.1", "lblOutName")
        .Caption = "3. Output File:": .Left = MARGIN: .Top = currentTop + 5: .Width = LBL_W: .Font.Bold = True
    End With
    
    ' Text Box (Locked)
    Set txtOutputName = Me.Controls.Add("Forms.TextBox.1", "txtOutputName")
    With txtOutputName
        .Left = MARGIN + LBL_W: .Top = currentTop: .Width = 300: .Height = CTRL_H
        .Text = "Click Browse to select path..."
        .Locked = True ' Read-only
        .BackColor = &HE0E0E0 ' Greyed out look
    End With
    
    ' Browse Button
    Set btnBrowse = Me.Controls.Add("Forms.CommandButton.1", "btnBrowse")
    With btnBrowse
        .Caption = "Browse...": .Left = MARGIN + LBL_W + 310: .Top = currentTop: .Width = 80: .Height = CTRL_H
    End With
    
    currentTop = currentTop + 35
    
    ' ============================
    ' SECTION 4: ACTIONS (New Line)
    ' ============================
    ' Center these buttons or place them comfortably below Section 3
    
    Set btnRun = Me.Controls.Add("Forms.CommandButton.1", "btnRun")
    With btnRun
        .Caption = "Generate JSON": .Left = MARGIN: .Top = currentTop: .Width = 140: .Height = 30: .Font.Bold = True: .Enabled = False
    End With
    
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close": .Left = MARGIN + 150: .Top = currentTop: .Width = 80: .Height = 30
    End With
    
    Me.Height = currentTop + 70
End Sub

' --- EVENT HANDLERS ---

Private Sub btnValidate_Click()
    Dim rng As Range
    Dim headers As Variant
    Dim firstDataRow As Variant
    Dim i As Long
    Dim cellVal As Variant
    Dim guessType As String
    
    On Error Resume Next
    Set rng = Range(refRangeSource.Text)
    On Error GoTo 0
    
    ' Basic Validation
    If rng Is Nothing Then MsgBox "Invalid range.", vbCritical: Exit Sub
    If rng.Rows.count < 2 Then MsgBox "Range must have header and data (at least 2 rows).", vbCritical: Exit Sub
    
    ' Load Headers (Row 1) and First Data Row (Row 2)
    headers = rng.Rows(1).Value
    firstDataRow = rng.Rows(2).Value
    
    ' Lock inputs, enable config frame
    ToggleInputs False
    btnReset.Enabled = True
    frameConfig.Enabled = True
    btnRun.Enabled = True
    
    lstColumns.Clear
    
    ' Loop through columns
    For i = 1 To UBound(headers, 2)
        lstColumns.AddItem headers(1, i)
        
        ' --- Auto-Detection Logic ---
        ' Get value from the first data row for this column
        cellVal = firstDataRow(1, i)
        
        ' Logic: If Numeric and not empty -> Aggregation, Else -> Category
        If IsNumeric(cellVal) And Not IsEmpty(cellVal) Then
            guessType = "aggregation"
        Else
            guessType = "category"
        End If
        
        lstColumns.List(i - 1, 1) = guessType
        lstColumns.List(i - 1, 2) = ""
    Next i
End Sub

Private Sub btnReset_Click()
    ToggleInputs True
    btnReset.Enabled = False
    frameConfig.Enabled = False
    lstColumns.Clear
    btnRun.Enabled = False
    txtOutputName.Text = "Click Browse to select path..."
    txtOutputName.BackColor = &HE0E0E0
End Sub

' --- BROWSE HANDLER ---
Private Sub btnBrowse_Click()
    Dim fName As Variant
    ' Open Save Dialog
    fName = Application.GetSaveAsFilename( _
        InitialFileName:="output_data.json", _
        FileFilter:="JSON Files (*.json), *.json", _
        Title:="Select Save Location")
        
    If fName <> False Then
        txtOutputName.Text = fName
        txtOutputName.BackColor = &HFFFFFF ' White background to indicate selection
    End If
End Sub

' --- CONFIG HANDLERS ---
Private Sub btnSetCate_Click()
    UpdateColumnStatus "category"
End Sub
Private Sub btnSetAgg_Click()
    UpdateColumnStatus "aggregation"
End Sub
Private Sub btnSetFormat_Click()
    Dim strFormat As String
    Dim i As Long, hasSel As Boolean
    For i = 0 To lstColumns.listCount - 1
        If lstColumns.Selected(i) Then hasSel = True: Exit For
    Next i
    If Not hasSel Then Exit Sub
    strFormat = InputBox("Format string (e.g. 0.00):", "Set Format", "0.00")
    If StrPtr(strFormat) = 0 Then Exit Sub
    For i = 0 To lstColumns.listCount - 1
        If lstColumns.Selected(i) Then
            lstColumns.List(i, 2) = strFormat
            lstColumns.Selected(i) = False
        End If
    Next i
End Sub

' --- RUN HANDLER ---
Private Sub btnRun_Click()
    Dim rng As Range
    Dim typeDict As Object, fmtDict As Object
    Dim jsonStr As String
    Dim i As Long, colName As String, colType As String, colFmt As String
    Dim hasCate As Boolean
    Dim filePath As String
    
    ' Check File Path
    filePath = txtOutputName.Text
    If InStr(filePath, ":") = 0 Or InStr(filePath, ".") = 0 Then
        MsgBox "Please select a valid output file path using 'Browse'.", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set rng = Range(refRangeSource.Text)
    On Error GoTo 0
    
    Set typeDict = CreateObject("Scripting.Dictionary")
    Set fmtDict = CreateObject("Scripting.Dictionary")
    
    hasCate = False
    For i = 0 To lstColumns.listCount - 1
        colName = lstColumns.List(i, 0)
        colType = lstColumns.List(i, 1)
        colFmt = lstColumns.List(i, 2)
        typeDict.item(colName) = colType
        If colType = "category" Then hasCate = True
        If Len(colFmt) > 0 Then fmtDict.item(colName) = colFmt
    Next i
    
    If Not hasCate Then
        MsgBox "Error: At least one column must be set as 'category'.", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    jsonStr = RangeToCompactJson(rng, typeDict, fmtDict)
    
'    Write to File
'    Dim fso As Object, ts As Object
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set ts = fso.CreateTextFile(filePath, True)
'    ts.Write jsonStr
'    ts.Close

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText jsonStr
        .SaveToFile filePath, 2 ' 2 = adSaveCreateOverWrite
        .Close
    End With
    
    MsgBox "JSON generated successfully!", vbInformation
    Unload Me
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' --- HELPERS ---
Private Sub UpdateColumnStatus(newType As String)
    Dim i As Long
    For i = 0 To lstColumns.listCount - 1
        If lstColumns.Selected(i) Then
            lstColumns.List(i, 1) = newType
            lstColumns.Selected(i) = False
        End If
    Next i
End Sub

Private Sub ToggleInputs(st As Boolean)
    If Not refRangeSource Is Nothing Then refRangeSource.Enabled = st
    btnValidate.Enabled = st
End Sub

