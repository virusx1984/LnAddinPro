Attribute VB_Name = "frm_format_mini"
Attribute VB_Base = "0{2C1E4298-FD44-4FDB-8D11-5553C357DE55}{C2058677-FF77-4968-A863-008D4B1C15FA}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' UserForm: frm_format_mini
' Purpose: A tiny dialog to ask for Header Row count. Default is 1.
Option Explicit

' Define controls with events
Private WithEvents lblPrompt As MSForms.Label
Attribute lblPrompt.VB_VarHelpID = -1
Private WithEvents txtRows As MSForms.TextBox
Attribute txtRows.VB_VarHelpID = -1
Private WithEvents btnRun As MSForms.CommandButton
Attribute btnRun.VB_VarHelpID = -1

' ==============================================================================
' INITIALIZATION
' ==============================================================================
Private Sub UserForm_Initialize()
    ' 1. Form Settings
    Me.Caption = "Format Table"
    Me.Width = 180
    Me.Height = 100
    Me.StartUpPosition = 1 ' CenterOwner
    
    ' 2. Label
    Set lblPrompt = Me.Controls.Add("Forms.Label.1", "lblPrompt")
    With lblPrompt
        .Caption = "Header Rows:"
        .Left = 15: .Top = 15: .Width = 80
        .Font.Size = 10
    End With
    
    ' 3. TextBox
    Set txtRows = Me.Controls.Add("Forms.TextBox.1", "txtRows")
    With txtRows
        .Text = "1" ' Default value
        .Left = 100: .Top = 12: .Width = 40: .Height = 18
        .TextAlign = fmTextAlignCenter
    End With
    
    ' 4. Button
    Set btnRun = Me.Controls.Add("Forms.CommandButton.1", "btnRun")
    With btnRun
        .Caption = "Format"
        .Left = 40: .Top = 40: .Width = 90: .Height = 24
        .Default = True ' Allows pressing Enter to trigger
        .BackColor = &H80FFFF ' Light Yellow style
    End With
End Sub

' ==============================================================================
' EVENT: ACTIVATE (Optimization #1: Auto-Select Text)
' ==============================================================================
Private Sub UserForm_Activate()
    ' Set focus and select all text so user can type immediately
    On Error Resume Next
    With txtRows
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    On Error GoTo 0
End Sub

' ==============================================================================
' LOGIC: BUTTON CLICK
' ==============================================================================
Private Sub btnRun_Click()
    Dim n As Long
    
    ' Validation
    If Not IsNumeric(txtRows.Text) Then
        MsgBox "Please enter a valid number.", vbExclamation
        Exit Sub
    End If
    
    n = CLng(txtRows.Text)
    If n < 0 Then n = 1
    
    ' Call the main logic (Assuming mod_format is the module name)
    ' Using the new name: ApplyTableStyle
    Call ApplyTableStyle(n)
    
    Unload Me
End Sub
