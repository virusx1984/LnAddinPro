' UserForm: frm_link_checker
' Purpose: Dynamic UI for filtering and viewing link audit results.
Option Explicit

' --- Controls with Events ---
' Filter Buttons
Public WithEvents btnAll As MSForms.CommandButton
Public WithEvents btnExt As MSForms.CommandButton
Public WithEvents btnLNF As MSForms.CommandButton
Public WithEvents btnInt As MSForms.CommandButton

' Main ListBox (Standard ListBox events are tricky with WithEvents,
' usually require a Class wrapper, but Click often works if declared properly
' or we use the MouseUp event if Click fails in dynamic context.
' For simplicity in this structure, we use MSForms.ListBox directly)
Public WithEvents lstResults As MSForms.ListBox

' Footer
Public WithEvents btnClose As MSForms.CommandButton

' Static Controls (No events needed logic-wise)
Private lblCount As MSForms.Label

' --- Data Storage ---
Private m_RawData As Variant ' Full dataset

' --- Layout Constants ---
Const MARGIN As Long = 10
Const BTN_H As Long = 24
Const GAP As Long = 5

' ==============================================================================
' INITIALIZATION & LAYOUT
' ==============================================================================
Private Sub UserForm_Initialize()
    Me.Caption = "Link & Function Checker"
    Me.Width = 500
    Me.Height = 400
    
    Dim currentTop As Long: currentTop = MARGIN
    Dim btnWidth As Long
    
    ' 1. Filter Buttons (Top Row)
    ' Calculate width for 4 buttons to fit equally
    btnWidth = (Me.InsideWidth - (MARGIN * 2) - (GAP * 3)) / 4
    
    Set btnAll = Me.Controls.Add("Forms.CommandButton.1", "btnAll")
    With btnAll
        .Caption = "All": .Left = MARGIN: .Top = currentTop
        .Width = btnWidth: .Height = BTN_H
        .BackColor = &H80FFFF ' Default Active
    End With
    
    Set btnExt = Me.Controls.Add("Forms.CommandButton.1", "btnExt")
    With btnExt
        .Caption = "External": .Left = MARGIN + btnWidth + GAP: .Top = currentTop
        .Width = btnWidth: .Height = BTN_H
    End With
    
    Set btnLNF = Me.Controls.Add("Forms.CommandButton.1", "btnLNF")
    With btnLNF
        .Caption = "LNF_Func": .Left = MARGIN + (btnWidth + GAP) * 2: .Top = currentTop
        .Width = btnWidth: .Height = BTN_H
    End With
    
    Set btnInt = Me.Controls.Add("Forms.CommandButton.1", "btnInt")
    With btnInt
        .Caption = "Internal": .Left = MARGIN + (btnWidth + GAP) * 3: .Top = currentTop
        .Width = btnWidth: .Height = BTN_H
    End With
    
    currentTop = currentTop + BTN_H + GAP + 5
    
    ' 2. ListBox Header Labels (Simulated)
    Dim col1W As Long: col1W = 50
    Dim col2W As Long: col2W = 60
    ' Remaining width for formula
    
    With Me.Controls.Add("Forms.Label.1", "lblH1")
        .Caption = "Address": .Left = MARGIN: .Top = currentTop: .Width = col1W: .Font.Bold = True
    End With
    With Me.Controls.Add("Forms.Label.1", "lblH2")
        .Caption = "Type": .Left = MARGIN + col1W: .Top = currentTop: .Width = col2W: .Font.Bold = True
    End With
    With Me.Controls.Add("Forms.Label.1", "lblH3")
        .Caption = "Formula": .Left = MARGIN + col1W + col2W: .Top = currentTop: .Width = 200: .Font.Bold = True
    End With
    
    currentTop = currentTop + 12 + GAP
    
    ' 3. Result ListBox
    Set lstResults = Me.Controls.Add("Forms.ListBox.1", "lstResults")
    With lstResults
        .Left = MARGIN
        .Top = currentTop
        .Width = Me.InsideWidth - (MARGIN * 2)
        .Height = Me.InsideHeight - currentTop - BTN_H - MARGIN - 20
        .ColumnCount = 3
        .ColumnWidths = CStr(col1W) & ";" & CStr(col2W) & ";" & CStr(Me.InsideWidth - col1W - col2W - 30)
        .Font.Name = "Segoe UI"
        .Font.Size = 9
    End With
    
    ' 4. Footer (Count Label + Close Button)
    Dim footerTop As Long
    footerTop = lstResults.Top + lstResults.Height + GAP
    
    Set lblCount = Me.Controls.Add("Forms.Label.1", "lblCount")
    With lblCount
        .Caption = "Ready": .Left = MARGIN: .Top = footerTop + 5: .Width = 200
        .ForeColor = &H808080
    End With
    
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close": .Left = Me.InsideWidth - 90: .Top = footerTop: .Width = 80: .Height = BTN_H
    End With
End Sub

' ==============================================================================
' PUBLIC METHODS
' ==============================================================================
Public Sub LoadData(data As Variant)
    m_RawData = data
    FilterList "All"
End Sub

' ==============================================================================
' LOGIC HANDLERS
' ==============================================================================

' --- Filter Buttons ---
Private Sub btnAll_Click()
    HighlightButton btnAll
    FilterList "All"
End Sub

Private Sub btnExt_Click()
    HighlightButton btnExt
    FilterList "External"
End Sub

Private Sub btnLNF_Click()
    HighlightButton btnLNF
    FilterList "LNF_Func"
End Sub

Private Sub btnInt_Click()
    HighlightButton btnInt
    FilterList "Internal"
End Sub

' Helper to filter data
Private Sub FilterList(category As String)
    Dim i As Long
    Dim count As Long
    
    lstResults.Clear
    If IsEmpty(m_RawData) Then Exit Sub
    
    count = 0
    For i = 1 To UBound(m_RawData, 1)
        Dim rowCat As String
        rowCat = m_RawData(i, 2)
        
        If category = "All" Or rowCat = category Then
            lstResults.AddItem m_RawData(i, 1)          ' Col 1: Address
            lstResults.List(lstResults.listCount - 1, 1) = rowCat       ' Col 2: Type
            lstResults.List(lstResults.listCount - 1, 2) = m_RawData(i, 3) ' Col 3: Formula
            count = count + 1
        End If
    Next i
    
    lblCount.Caption = "Items Found: " & count
End Sub

' Helper to change button colors
Private Sub HighlightButton(activeBtn As MSForms.CommandButton)
    ' Reset all to system color
    btnAll.BackColor = &H8000000F
    btnExt.BackColor = &H8000000F
    btnLNF.BackColor = &H8000000F
    btnInt.BackColor = &H8000000F
    
    ' Highlight active
    activeBtn.BackColor = &H80FFFF ' Light Yellow
End Sub

' --- ListBox Interaction ---
Private Sub lstResults_Click()
    Dim addr As String
    If lstResults.ListIndex = -1 Then Exit Sub
    
    addr = lstResults.List(lstResults.ListIndex, 0)
    
    On Error Resume Next
    ActiveSheet.Range(addr).Select
    On Error GoTo 0
End Sub

' --- Footer ---
Private Sub btnClose_Click()
    Unload Me
End Sub