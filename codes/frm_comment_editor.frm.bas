' UserForm: frm_comment_editor
' Purpose: A dynamic central editor for Excel cell comments (Notes).
' Features: Supports TAB character input within text, Dynamic UI, Persistent window, Quick Delete.
Option Explicit

' --- Event Handlers for Dynamic Controls ---
Public WithEvents btnSave As MSForms.CommandButton
Public WithEvents btnDelete As MSForms.CommandButton ' NEW: Delete button
Public WithEvents btnClose As MSForms.CommandButton

' --- Control References ---
Private lblTarget As MSForms.Label
Private txtComment As MSForms.TextBox
Private lblStatus As MSForms.Label

' --- Data Variables ---
Private m_TargetCell As Range

' ==============================================================================
' UI INITIALIZATION & DATA LOADING
' ==============================================================================
Private Sub UserForm_Initialize()
    ' 1. Validate Selection
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell first.", vbExclamation, "Invalid Selection"
        Application.OnTime Now, "CloseCommentEditor"
        Exit Sub
    End If
    
    Set m_TargetCell = ActiveCell
    
    ' 2. Form Properties
    Me.Caption = "Comment Editor"
    Me.Width = 380
    Me.Height = 330
    
    Dim currentTop As Single: currentTop = 10
    Const MARGIN As Single = 10
    Const CTRL_W As Single = 340
    
    ' --- Section 1: Target Cell Info ---
    Set lblTarget = Me.Controls.Add("Forms.Label.1", "lblTarget")
    With lblTarget
        .Caption = "Editing Comment for Cell: " & m_TargetCell.Address(False, False)
        .Left = MARGIN: .Top = currentTop: .Width = CTRL_W
        .Font.bold = True: .ForeColor = &H800000
    End With
    currentTop = currentTop + 20
    
    ' --- Section 2: Comment Text Area ---
    Set txtComment = Me.Controls.Add("Forms.TextBox.1", "txtComment")
    With txtComment
        .Left = MARGIN: .Top = currentTop: .Width = CTRL_W: .Height = 180
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .EnterKeyBehavior = True
        .TabKeyBehavior = True ' Allows TAB to insert tab indentations
        .Font.Name = "Segoe UI": .Font.Size = 10
        .TabIndex = 0
    End With
    currentTop = currentTop + 190
    
    ' --- Section 3: Action Buttons ---
    Set btnSave = Me.Controls.Add("Forms.CommandButton.1", "btnSave")
    With btnSave
        .Caption = "Save Note"
        .Left = MARGIN: .Top = currentTop: .Width = 100: .Height = 28
        .BackColor = &H80FF80: .Font.bold = True
        .TabIndex = 1
    End With
    
    ' NEW: Delete Button
    Set btnDelete = Me.Controls.Add("Forms.CommandButton.1", "btnDelete")
    With btnDelete
        .Caption = "Delete Note"
        .Left = MARGIN + 110: .Top = currentTop: .Width = 90: .Height = 28
        .BackColor = &H8080FF ' Light Red
        .TabIndex = 2
    End With
    
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close"
        .Left = Me.InsideWidth - 90: .Top = currentTop: .Width = 80: .Height = 28
        .TabIndex = 3
    End With
    currentTop = currentTop + 35
    
    ' --- Section 4: Status Feedback ---
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus")
    With lblStatus
        .Left = MARGIN: .Top = currentTop: .Width = CTRL_W
        .ForeColor = &H8000&: .Caption = ""
    End With
    
    ' 3. Load Existing Comment
    LoadExistingComment
End Sub

' ==============================================================================
' LOGIC: DATA HANDLING
' ==============================================================================
Private Sub LoadExistingComment()
    If Not m_TargetCell Is Nothing Then
        If Not m_TargetCell.Comment Is Nothing Then
            txtComment.Text = m_TargetCell.Comment.Text
            txtComment.SelStart = Len(txtComment.Text)
        Else
            txtComment.Text = ""
        End If
    End If
End Sub

' ==============================================================================
' LOGIC: ACTION BUTTONS
' ==============================================================================
Private Sub btnSave_Click()
    Dim rawText As String
    Dim renderText As String
    
    rawText = txtComment.Text
    ' Fix Excel limitation: convert tabs to spaces
    renderText = Replace(rawText, vbTab, "    ")
    
    On Error GoTo ErrorHandler
    
    lblStatus.ForeColor = &H8000&
    
    If Len(Trim(renderText)) = 0 Then
        If Not m_TargetCell.Comment Is Nothing Then
            m_TargetCell.Comment.Delete
            lblStatus.Caption = "Note deleted successfully at " & Format(Now, "hh:mm:ss")
        End If
    Else
        If m_TargetCell.Comment Is Nothing Then m_TargetCell.AddComment ""
        
        m_TargetCell.Comment.Text Text:=renderText
        m_TargetCell.Comment.Shape.TextFrame.AutoSize = True
        
        lblStatus.Caption = "Note saved successfully at " & Format(Now, "hh:mm:ss")
    End If
    
    txtComment.SetFocus
    Exit Sub
    
ErrorHandler:
    lblStatus.ForeColor = &HFF&
    lblStatus.Caption = "Error: Could not save note."
End Sub

' NEW: Direct Delete Logic
Private Sub btnDelete_Click()
    On Error GoTo ErrorHandler
    
    If Not m_TargetCell Is Nothing Then
        If Not m_TargetCell.Comment Is Nothing Then
            m_TargetCell.Comment.Delete
        End If
    End If
    
    ' Close the form immediately after direct deletion
    Unload Me
    Exit Sub
    
ErrorHandler:
    MsgBox "Failed to delete comment. The sheet might be protected.", vbCritical, "Error"
    Unload Me
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub