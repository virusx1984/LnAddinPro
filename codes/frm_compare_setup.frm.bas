Attribute VB_Name = "frm_compare_setup"
Attribute VB_Base = "0{CC954C11-04D7-49A1-A693-9B5307FA9CFD}{31E73844-3361-49A3-BCBF-98A78FE39215}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' UserForm: frm_compare_setup
' Purpose: Compare Setup with Index, Ignore, and Reference Source (A vs B) selection.
'          (Simplified: No Re-ordering)
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

' Reference Buttons
Public WithEvents btnRefUseA As MSForms.CommandButton ' Reference from Range A
Attribute btnRefUseA.VB_VarHelpID = -1
Public WithEvents btnRefUseB As MSForms.CommandButton ' Reference from Range B
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

' --- 3. Layout Constants ---
Const MARGIN As Long = 10
Const CTRL_H As Long = 20
Const GAP As Long = 5
Const LBL_W As Long = 80
Const INPUT_W As Long = 180

Private Sub UserForm_Initialize()
    
    Me.Caption = "Compare Setup"
    Me.Width = 480
    Me.Height = 450 ' Height reduced since we removed buttons
    
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
        .Height = 230 ' Height reduced
        .Enabled = False
    End With
    
    ' ListBox
    Set lstColumns = frameConfig.Controls.Add("Forms.ListBox.1", "lstColumns")
    With lstColumns
        .Left = 10: .Top = 20: .Width = 260: .Height = 200
        .ColumnCount = 2: .ColumnWidths = "160;80"
        .MultiSelect = fmMultiSelectExtended
    End With
    
    ' --- Buttons (Right Side) ---
    Dim btnLeft As Long: btnLeft = 280
    Dim btnTop As Long: btnTop = 20
    
    ' Label: Key (Corrected syntax)
    With frameConfig.Controls.Add("Forms.Label.1", "lblKey")
        .Caption = "Match By:"
        .Left = btnLeft
        .Top = btnTop
        .Width = 80
        .Height = 15
    End With
    btnTop = btnTop + 15
    
    Set btnSetIndex = frameConfig.Controls.Add("Forms.CommandButton.1", "btnSetIndex")
    With btnSetIndex
        .Caption = "INDEX (Key)": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
    End With
    btnTop = btnTop + 30
    
    ' Label: Reference (Corrected syntax)
    With frameConfig.Controls.Add("Forms.Label.1", "lblRef")
        .Caption = "Reference Source:"
        .Left = btnLeft
        .Top = btnTop
        .Width = 100
        .Height = 15
    End With
    btnTop = btnTop + 15
    
    ' REF A Button
    Set btnRefUseA = frameConfig.Controls.Add("Forms.CommandButton.1", "btnRefUseA")
    With btnRefUseA
        .Caption = "REF (Use A)": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
        .ControlTipText = "Reference column: Value taken from Range A"
    End With
    btnTop = btnTop + 25
    
    ' REF B Button
    Set btnRefUseB = frameConfig.Controls.Add("Forms.CommandButton.1", "btnRefUseB")
    With btnRefUseB
        .Caption = "REF (Use B)": .Left = btnLeft: .Top = btnTop: .Width = 100: .Height = 22
        .ControlTipText = "Reference column: Value taken from Range B"
    End With
    btnTop = btnTop + 30
    
    ' Label: Others (Corrected syntax)
    With frameConfig.Controls.Add("Forms.Label.1", "lblOth")
        .Caption = "Comparison:"
        .Left = btnLeft
        .Top = btnTop
        .Width = 80
        .Height = 15
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
    ' SECTION 3: ACTION
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
    
    On Error Resume Next
    Set rng1 = Range(refRange1.Text)
    Set rng2 = Range(refRange2.Text)
    On Error GoTo 0
    
    If rng1 Is Nothing Or rng2 Is Nothing Then MsgBox "Invalid ranges.", vbCritical: Exit Sub
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
Private Sub btnRun_Click()
    Dim indexCols As String
    Dim ignoreCols As String
    Dim referenceCols As String
    Dim vlookupOrderCols As String ' Columns that need to pull from Range B
    
    Dim i As Long
    Dim status As String
    Dim colName As String
    
    ' Loop through the ListBox
    For i = 0 To lstColumns.listCount - 1
        colName = lstColumns.List(i, 0)
        status = lstColumns.List(i, 1)
        
        Select Case status
            Case "INDEX"
                indexCols = indexCols & colName & ","
                
            Case "IGNORE"
                ignoreCols = ignoreCols & colName & ","
                
            Case "REF: Range A"
                referenceCols = referenceCols & colName & ","
                
            Case "REF: Range B"
                referenceCols = referenceCols & colName & ","
                vlookupOrderCols = vlookupOrderCols & colName & ","
                
            ' Case "COMPARE" is implicit
        End Select
    Next i
    
    ' Clean commas
    indexCols = StripComma(indexCols)
    ignoreCols = StripComma(ignoreCols)
    referenceCols = StripComma(referenceCols)
    vlookupOrderCols = StripComma(vlookupOrderCols)
    
    If Len(indexCols) = 0 Then
        MsgBox "Please select at least one INDEX column.", vbExclamation
        Exit Sub
    End If
    
    ' --- DEBUG: Show Parameters ---
    Dim msg As String
    msg = "Execute Parameters:" & vbCrLf & _
          "Ranges: " & refRange1.Text & " vs " & refRange2.Text & vbCrLf & _
          "------------------------" & vbCrLf & _
          "INDEX: " & indexCols & vbCrLf & _
          "IGNORE: " & ignoreCols & vbCrLf & _
          "REFERENCE: " & referenceCols & vbCrLf & _
          "USE RANGE B FOR: " & vlookupOrderCols
          
    MsgBox msg, vbInformation, "Debug"
    
    ' Call your actual function here:
    ' YourFunction Range(refRange1.Text), Range(refRange2.Text), indexCols, ignoreCols, referenceCols, vlookupOrderCols
    
    Unload Me
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

