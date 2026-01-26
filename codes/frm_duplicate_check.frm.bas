Attribute VB_Name = "frm_duplicate_check"
Attribute VB_Base = "0{5A93E12F-7B1B-4F61-8C08-70CC263868B4}{F072B9F0-41B7-4D6F-A0B7-4F58E8BBFF9F}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' UserForm: frm_duplicate_check
Option Explicit

' --- Event Handlers for Dynamic Controls ---
Public WithEvents btnAnalyze As MSForms.CommandButton
Attribute btnAnalyze.VB_VarHelpID = -1
Public WithEvents btnHighlight As MSForms.CommandButton
Attribute btnHighlight.VB_VarHelpID = -1
Public WithEvents btnSelect As MSForms.CommandButton
Attribute btnSelect.VB_VarHelpID = -1
Public WithEvents btnClose As MSForms.CommandButton
Attribute btnClose.VB_VarHelpID = -1

' --- Control References ---
Private refRange As Object          ' RefEdit or TextBox
Private lstColumns As MSForms.ListBox
Private lblStatus As MSForms.Label
Private lblDetail As MSForms.Label
Private frameAction As MSForms.Frame

' --- Data Variables ---
Private m_TargetRange As Range
Private m_DuplicateRanges As Range  ' Stores the result range (union of duplicates)
Private Const CELL_LIMIT As Long = 100000 ' Conservative Threshold

' ==============================================================================
' UI INITIALIZATION
' ==============================================================================
Private Sub UserForm_Initialize()
    Me.Caption = "Duplicate Check Wizard"
    Me.Width = 360
    ' Reduced height since we removed the listbox
    Me.Height = 320
    
    Dim currentTop As Single: currentTop = 10
    Const MARGIN As Single = 10
    Const CTRL_W As Single = 320
    
    ' --- Section 1: Range Selection ---
    With Me.Controls.Add("Forms.Label.1", "lblRng")
        .Caption = "1. Select Data Range:"
        .Left = MARGIN: .Top = currentTop: .Width = CTRL_W
    End With
    currentTop = currentTop + 15
    
    ' Try to create RefEdit, fallback to TextBox
    On Error Resume Next
    Set refRange = Me.Controls.Add("RefEdit.Ctrl", "refRange")
    If Err.Number <> 0 Then Set refRange = Me.Controls.Add("Forms.TextBox.1", "refRange")
    On Error GoTo 0
    
    With refRange
        .Left = MARGIN: .Top = currentTop: .Width = CTRL_W: .Height = 18
        If TypeName(Selection) = "Range" Then
            If Selection.Cells.CountLarge <= CELL_LIMIT Then
                .Value = Selection.Address(External:=False)
            End If
        End If
    End With
    currentTop = currentTop + 35
    
    ' --- Section 2: Analysis (Moved up) ---
    Set btnAnalyze = Me.Controls.Add("Forms.CommandButton.1", "btnAnalyze")
    With btnAnalyze
        .Caption = "Analyze Cells"
        .Left = MARGIN: .Top = currentTop: .Width = 120: .Height = 28
        .BackColor = &H80FFFF ' Light Yellow
        .Font.Bold = True
    End With
    
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close"
        .Left = Me.InsideWidth - 90: .Top = currentTop: .Width = 80: .Height = 28
    End With
    currentTop = currentTop + 40
    
    ' --- Section 3: Results & Actions (Hidden initially) ---
    Set frameAction = Me.Controls.Add("Forms.Frame.1", "frameAction")
    With frameAction
        .Caption = "Results & Actions"
        .Left = MARGIN: .Top = currentTop: .Width = CTRL_W: .Height = 100
        .Visible = False ' Hidden until analyzed
    End With
    
    ' Status Labels inside Frame
    Set lblStatus = frameAction.Controls.Add("Forms.Label.1", "lblStatus")
    With lblStatus
        .Caption = "Status: Waiting..."
        .Left = 10: .Top = 15: .Width = 300: .Font.Bold = True
    End With
    
    Set lblDetail = frameAction.Controls.Add("Forms.Label.1", "lblDetail")
    With lblDetail
        .Caption = "Details..."
        .Left = 10: .Top = 30: .Width = 300: .ForeColor = &H808080
    End With
    
    ' Action Buttons inside Frame
    Set btnHighlight = frameAction.Controls.Add("Forms.CommandButton.1", "btnHighlight")
    With btnHighlight
        .Caption = "Highlight All"
        .Left = 10: .Top = 60: .Width = 90: .Height = 24
        .BackColor = &H8080FF ' Light Red
    End With
    
    Set btnSelect = frameAction.Controls.Add("Forms.CommandButton.1", "btnSelect")
    With btnSelect
        .Caption = "Select Cells"
        .Left = 110: .Top = 60: .Width = 90: .Height = 24
    End With
    
    ' Adjust form height based on final position
    Me.Height = currentTop + 140
End Sub

' ==============================================================================
' LOGIC: Helper to populate columns
' ==============================================================================
Private Sub PopulateColumns()
    Dim rng As Range
    Dim i As Long
    
    On Error Resume Next
    Set rng = Range(refRange.Value)
    On Error GoTo 0
    
    lstColumns.Clear
    If rng Is Nothing Then Exit Sub
    
    ' Only read the header row
    Dim headerArr As Variant
    headerArr = rng.Rows(1).Value
    
    If IsArray(headerArr) Then
        For i = 1 To UBound(headerArr, 2)
            lstColumns.AddItem headerArr(1, i)
            lstColumns.Selected(i - 1) = True ' Default select all
        Next i
    Else
        ' Single cell case
        lstColumns.AddItem rng.Rows(1).Value
        lstColumns.Selected(0) = True
    End If
End Sub

' ==============================================================================
' LOGIC: ANALYZE BUTTON (Single Cell Mode - Highlight ALL Duplicates)
' ==============================================================================
Private Sub btnAnalyze_Click()
    Dim r As Long, c As Long
    Dim cellVal As String
    Dim dict As Object
    Dim dataArr As Variant
    Dim dupCount As Long
    
    ' 1. Validate Range
    On Error Resume Next
    Set m_TargetRange = Range(refRange.Value)
    On Error GoTo 0
    
    If m_TargetRange Is Nothing Then
        MsgBox "Invalid range selected.", vbExclamation
        Exit Sub
    End If
    
    ' 2. SAFETY CHECK
    If m_TargetRange.Cells.CountLarge > CELL_LIMIT Then
        MsgBox "Safety Stop: Selection exceeds " & Format(CELL_LIMIT, "#,##0") & " cells.", vbCritical
        Exit Sub
    End If
    
    ' 3. Perform Analysis
    Set dict = CreateObject("Scripting.Dictionary")
    Set m_DuplicateRanges = Nothing
    dupCount = 0
    
    ' Read to array for speed
    ' Handle single cell selection edge case to prevent error
    If m_TargetRange.Cells.CountLarge = 1 Then
        ReDim dataArr(1 To 1, 1 To 1)
        dataArr(1, 1) = m_TargetRange.Value
    Else
        dataArr = m_TargetRange.Value
    End If
    
    ' --- STEP 1: Count Frequencies ---
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            cellVal = Trim(CStr(dataArr(r, c)))
            
            If cellVal <> "" Then ' Skip empty cells
                If dict.Exists(cellVal) Then
                    dict(cellVal) = dict(cellVal) + 1
                Else
                    dict.Add cellVal, 1
                End If
            End If
        Next c
    Next r
    
    ' --- STEP 2: Collect ALL cells involved in duplication ---
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            cellVal = Trim(CStr(dataArr(r, c)))
            
            If cellVal <> "" Then
                If dict.Exists(cellVal) Then
                    ' If the total count is > 1, this cell is part of a duplicate set
                    If dict(cellVal) > 1 Then
                        dupCount = dupCount + 1
                        
                        If m_DuplicateRanges Is Nothing Then
                            Set m_DuplicateRanges = m_TargetRange.Cells(r, c)
                        Else
                            Set m_DuplicateRanges = Union(m_DuplicateRanges, m_TargetRange.Cells(r, c))
                        End If
                    End If
                End If
            End If
        Next c
    Next r
    
    ' 4. Update UI
    frameAction.Visible = True
    If dupCount > 0 Then
        lblStatus.Caption = "Status: Found Duplicates!"
        lblStatus.ForeColor = &HFF& ' Red
        lblDetail.Caption = "Total cells involved: " & dupCount
        btnHighlight.Enabled = True
        btnSelect.Enabled = True
    Else
        lblStatus.Caption = "Status: No Duplicates."
        lblStatus.ForeColor = &H8000& ' Green
        lblDetail.Caption = "All cells contain unique values."
        btnHighlight.Enabled = False
        btnSelect.Enabled = False
    End If
    
    ' Clean up
    Set dict = Nothing
End Sub

' ==============================================================================
' LOGIC: ACTION BUTTONS
' ==============================================================================
Private Sub btnHighlight_Click()
    If m_DuplicateRanges Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    ' Only highlight the specific duplicate cells
    m_DuplicateRanges.Interior.Color = vbRed
    Application.ScreenUpdating = True
    
    MsgBox "Duplicate cells highlighted in Red.", vbInformation
    Unload Me
End Sub

Private Sub btnSelect_Click()
    If m_DuplicateRanges Is Nothing Then Exit Sub
    ' Select only the duplicate cells
    m_DuplicateRanges.Select
    MsgBox "Duplicate cells selected.", vbInformation
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' Helper to refresh columns if user manually changes range text
Private Sub refRange_Change()
    ' Optional: Re-populate columns on change (debounce recommended in real apps)
    ' keeping simple for now
End Sub

