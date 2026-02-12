' UserForm: frm_link_checker
' Purpose: Dynamic UI for Link Audit (Optimized for Large Data).
Option Explicit

' --- Controls ---
Public WithEvents btnAll As MSForms.CommandButton
Public WithEvents btnExt As MSForms.CommandButton
Public WithEvents btnLNF As MSForms.CommandButton
Public WithEvents btnInt As MSForms.CommandButton

Public WithEvents lstResults As MSForms.ListBox
Public WithEvents btnClose As MSForms.CommandButton
Private lblCount As MSForms.Label
Private lblStatus As MSForms.Label ' New: Show Total vs Displayed

' --- Data ---
Private m_RawData As Variant    ' The subset of data for display (Max 2000 rows)
Private m_TotalFound As Long    ' The actual total count found in sheet

' --- Layout ---
Const MARGIN As Long = 10
Const BTN_H As Long = 24
Const GAP As Long = 5

Private Sub UserForm_Initialize()
    Me.Caption = "Link & Function Checker"
    Me.Width = 500
    Me.Height = 400
    
    DrawControls
End Sub

Private Sub DrawControls()
    Dim currentTop As Long: currentTop = MARGIN
    Dim btnWidth As Long
    
    ' 1. Filter Buttons
    btnWidth = (Me.InsideWidth - (MARGIN * 2) - (GAP * 3)) / 4
    
    Set btnAll = CreateBtn("btnAll", "All", MARGIN, currentTop, btnWidth)
    Set btnExt = CreateBtn("btnExt", "External", MARGIN + btnWidth + GAP, currentTop, btnWidth)
    Set btnLNF = CreateBtn("btnLNF", "LNF_Func", MARGIN + (btnWidth + GAP) * 2, currentTop, btnWidth)
    Set btnInt = CreateBtn("btnInt", "Internal", MARGIN + (btnWidth + GAP) * 3, currentTop, btnWidth)
    
    btnAll.BackColor = &H80FFFF ' Active default
    currentTop = currentTop + BTN_H + GAP + 5
    
    ' 2. Headers
    Dim col1W As Long: col1W = 60
    Dim col2W As Long: col2W = 60
    
    CreateLabel "Address", MARGIN, currentTop, col1W, True
    CreateLabel "Type", MARGIN + col1W, currentTop, col2W, True
    CreateLabel "Formula", MARGIN + col1W + col2W, currentTop, 200, True
    
    currentTop = currentTop + 12 + GAP
    
    ' 3. ListBox
    Set lstResults = Me.Controls.Add("Forms.ListBox.1", "lstResults")
    With lstResults
        .Left = MARGIN
        .Top = currentTop
        .Width = Me.InsideWidth - (MARGIN * 2)
        .Height = Me.InsideHeight - currentTop - BTN_H - MARGIN - 25
        .ColumnCount = 3
        .ColumnWidths = CStr(col1W) & ";" & CStr(col2W) & ";" & CStr(Me.InsideWidth - col1W - col2W - 30)
        .Font.Name = "Segoe UI"
        .Font.Size = 9
    End With
    
    ' 4. Footer
    Dim footerTop As Long: footerTop = lstResults.Top + lstResults.Height + GAP
    
    Set lblCount = Me.Controls.Add("Forms.Label.1", "lblCount")
    With lblCount
        .Caption = "Ready": .Left = MARGIN: .Top = footerTop: .Width = 300: .ForeColor = &H404040
    End With
    
    Set btnClose = CreateBtn("btnClose", "Close", Me.InsideWidth - 90, footerTop, 80)
End Sub

' Helper to create buttons
Private Function CreateBtn(n As String, c As String, l As Long, t As Long, w As Long) As MSForms.CommandButton
    Set CreateBtn = Me.Controls.Add("Forms.CommandButton.1", n)
    With CreateBtn
        .Caption = c: .Left = l: .Top = t: .Width = w: .Height = BTN_H
    End With
End Function

Private Sub CreateLabel(c As String, l As Long, t As Long, w As Long, bold As Boolean)
    With Me.Controls.Add("Forms.Label.1", "")
        .Caption = c: .Left = l: .Top = t: .Width = w
        If bold Then .Font.bold = True
    End With
End Sub

' ==============================================================================
' LOGIC
' ==============================================================================

Public Sub LoadData(data As Variant, totalCnt As Long)
    m_RawData = data
    m_TotalFound = totalCnt
    FilterList "All"
End Sub

' OPTIMIZED FILTER LOGIC: Use Array Assignment
Private Sub FilterList(category As String)
    Dim i As Long, cnt As Long
    Dim arrDisplay() As Variant
    Dim rawRows As Long
    
    ' Clear current list
    lstResults.Clear
    
    If IsEmpty(m_RawData) Then
        lblCount.Caption = "No items."
        Exit Sub
    End If
    
    rawRows = UBound(m_RawData, 1)
    
    ' 1. First Pass: Count matches to size the array
    cnt = 0
    For i = 1 To rawRows
        If category = "All" Or m_RawData(i, 2) = category Then
            cnt = cnt + 1
        End If
    Next i
    
    If cnt = 0 Then
        lblCount.Caption = "No items found for filter: " & category
        Exit Sub
    End If
    
    ' 2. Second Pass: Fill Display Array
    ' ListBox.List expects (0 to rows-1, 0 to cols-1)
    ReDim arrDisplay(0 To cnt - 1, 0 To 2)
    Dim currIdx As Long
    currIdx = 0
    
    For i = 1 To rawRows
        If category = "All" Or m_RawData(i, 2) = category Then
            arrDisplay(currIdx, 0) = m_RawData(i, 1) ' Addr
            arrDisplay(currIdx, 1) = m_RawData(i, 2) ' Type
            arrDisplay(currIdx, 2) = m_RawData(i, 3) ' Formula
            currIdx = currIdx + 1
        End If
    Next i
    
    ' 3. Bulk Assignment (Instant)
    lstResults.List = arrDisplay
    
    ' 4. Update Status Label
    Dim msg As String
    msg = "Showing " & cnt & " items."
    If m_TotalFound > rawRows Then
        msg = msg & " (Note: First " & rawRows & " of " & m_TotalFound & " total errors shown)"
    End If
    lblCount.Caption = msg
End Sub

' Button Handlers
Private Sub btnAll_Click(): ResetColors: btnAll.BackColor = &H80FFFF: FilterList "All": End Sub
Private Sub btnExt_Click(): ResetColors: btnExt.BackColor = &H80FFFF: FilterList "External": End Sub
Private Sub btnLNF_Click(): ResetColors: btnLNF.BackColor = &H80FFFF: FilterList "LNF_Func": End Sub
Private Sub btnInt_Click(): ResetColors: btnInt.BackColor = &H80FFFF: FilterList "Internal": End Sub

Private Sub ResetColors()
    btnAll.BackColor = &H8000000F: btnExt.BackColor = &H8000000F
    btnLNF.BackColor = &H8000000F: btnInt.BackColor = &H8000000F
End Sub

Private Sub lstResults_Click()
    If lstResults.ListIndex = -1 Then Exit Sub
    On Error Resume Next
    ActiveSheet.Range(lstResults.List(lstResults.ListIndex, 0)).Select
    On Error GoTo 0
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub