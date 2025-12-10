Attribute VB_Name = "frm_code_export"
Attribute VB_Base = "0{EB4A5704-085F-435C-BA9C-C8B10B09489D}{9D38C87B-6270-4D23-A8EE-76204F30F35D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' UserForm: frm_code_export
' Purpose: UI for selecting range and outputting Markdown/HTML code.
Option Explicit

' --- Controls ---
Public WithEvents btnRun As MSForms.CommandButton
Attribute btnRun.VB_VarHelpID = -1
Public WithEvents btnCopy As MSForms.CommandButton
Attribute btnCopy.VB_VarHelpID = -1
Public WithEvents btnClose As MSForms.CommandButton
Attribute btnClose.VB_VarHelpID = -1

Public WithEvents optMarkdown As MSForms.OptionButton
Attribute optMarkdown.VB_VarHelpID = -1
Public WithEvents optHTML As MSForms.OptionButton
Attribute optHTML.VB_VarHelpID = -1
Public WithEvents chkBootstrap As MSForms.CheckBox
Attribute chkBootstrap.VB_VarHelpID = -1

Public refRange As Object
Public txtResult As MSForms.TextBox

' --- Layout Constants ---
Const MARGIN As Long = 10
Const CTRL_H As Long = 20
Const GAP As Long = 5

Private Sub UserForm_Initialize()
    Me.Caption = "Export Range to Code"
    Me.Width = 450
    Me.Height = 400
    
    Dim currentTop As Long: currentTop = MARGIN
    
    ' 1. Range Selection
    With Me.Controls.Add("Forms.Label.1", "lblRng")
        .Caption = "Select Range:": .Left = MARGIN: .Top = currentTop: .Width = 80
    End With
    
    On Error Resume Next
    Set refRange = Me.Controls.Add("RefEdit.Ctrl", "refRange")
    If Err.Number <> 0 Then Set refRange = Me.Controls.Add("Forms.TextBox.1", "refRange")
    On Error GoTo 0
    With refRange
        .Left = MARGIN + 90: .Top = currentTop: .Width = 320: .Height = CTRL_H
        If TypeName(Selection) = "Range" Then .Text = Selection.Address(External:=False)
    End With
    currentTop = currentTop + CTRL_H + GAP + 5
    
    ' 2. Format Options (Radio Buttons)
    With Me.Controls.Add("Forms.Label.1", "lblFmt")
        .Caption = "Format:": .Left = MARGIN: .Top = currentTop: .Width = 80
    End With
    
    Set optMarkdown = Me.Controls.Add("Forms.OptionButton.1", "optMarkdown")
    With optMarkdown
        .Caption = "Markdown (MD)": .Left = MARGIN + 90: .Top = currentTop: .Width = 100: .Value = True
        .GroupName = "ExportFmt"
    End With
    
    Set optHTML = Me.Controls.Add("Forms.OptionButton.1", "optHTML")
    With optHTML
        .Caption = "HTML": .Left = MARGIN + 200: .Top = currentTop: .Width = 60
        .GroupName = "ExportFmt"
    End With
    
    Set chkBootstrap = Me.Controls.Add("Forms.CheckBox.1", "chkBootstrap")
    With chkBootstrap
        .Caption = "Add Bootstrap Class": .Left = MARGIN + 270: .Top = currentTop: .Width = 140
        .Enabled = False ' Disabled by default (MD selected)
    End With
    currentTop = currentTop + CTRL_H + GAP + 5
    
    ' 3. Action Buttons
    Set btnRun = Me.Controls.Add("Forms.CommandButton.1", "btnRun")
    With btnRun
        .Caption = "Generate Code": .Left = MARGIN: .Top = currentTop: .Width = 100: .Height = 24: .BackColor = &H80FF80
    End With
    
    Set btnCopy = Me.Controls.Add("Forms.CommandButton.1", "btnCopy")
    With btnCopy
        .Caption = "Copy to Clipboard": .Left = MARGIN + 110: .Top = currentTop: .Width = 120: .Height = 24
    End With
    
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close": .Left = Me.InsideWidth - 90: .Top = currentTop: .Width = 80: .Height = 24
    End With
    currentTop = currentTop + 30
    
    ' 4. Output Text Area
    With Me.Controls.Add("Forms.Label.1", "lblRes")
        .Caption = "Result:": .Left = MARGIN: .Top = currentTop: .Width = 80
    End With
    currentTop = currentTop + 15
    
    Set txtResult = Me.Controls.Add("Forms.TextBox.1", "txtResult")
    With txtResult
        .Left = MARGIN: .Top = currentTop: .Width = Me.InsideWidth - (MARGIN * 2)
        .Height = Me.InsideHeight - currentTop - MARGIN - 20
        .Multiline = True
        .ScrollBars = fmScrollBarsVertical
        .Font.Name = "Consolas" ' Monospace font for code
        .Font.Size = 9
    End With
End Sub

' --- Logic Handlers ---

' Toggle Bootstrap checkbox based on HTML selection
Private Sub optMarkdown_Click()
    If Not chkBootstrap Is Nothing Then
        chkBootstrap.Enabled = False
    End If
End Sub

Private Sub optHTML_Click()
    If Not chkBootstrap Is Nothing Then
        chkBootstrap.Enabled = True
        chkBootstrap.Value = True ' Default to on for HTML
    End If
End Sub

' Generate Code
Private Sub btnRun_Click()
    Dim rng As Range
    Dim resultStr As String
    
    On Error Resume Next
    Set rng = Range(refRange.Text)
    On Error GoTo 0
    
    If rng Is Nothing Then MsgBox "Invalid Range", vbCritical: Exit Sub
    
    If optMarkdown.Value Then
        resultStr = mod_funcs.RangeToMarkdown(rng)
    Else
        resultStr = mod_funcs.RangeToHTML(rng, chkBootstrap.Value)
    End If
    
    txtResult.Text = resultStr
End Sub

' Copy to Clipboard
Private Sub btnCopy_Click()
    If txtResult.Text = "" Then Exit Sub
    
    ' Fix: Use Early Binding instead of CreateObject to avoid Error 429
    Dim dataObj As MSForms.DataObject
    Set dataObj = New MSForms.DataObject
    
    On Error GoTo ErrHandler
    dataObj.SetText txtResult.Text
    dataObj.PutInClipboard
    MsgBox "Copied to clipboard!", vbInformation
    Exit Sub
    
ErrHandler:
    ' Fallback: If DataObject fails (common on some Windows versions), select text for manual copy
    MsgBox "Clipboard access failed. Text selected - please press Ctrl+C.", vbExclamation
    txtResult.SetFocus
    txtResult.SelStart = 0
    txtResult.SelLength = Len(txtResult.Text)
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
