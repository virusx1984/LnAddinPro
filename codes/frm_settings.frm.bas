Option Explicit

' --- Event Handlers for Dynamic Controls ---
' We declare this WithEvents so we can capture the Click event
Public WithEvents btnSave As MSForms.CommandButton

' --- Control References ---
' We keep references to these to access their properties easily later
Private txtApiEndpoint As MSForms.TextBox
Private lblEndpoint As MSForms.Label

' --- Data Variables ---
Private m_ConfigPath As String
Private m_JsonData As Scripting.Dictionary

Private Sub UserForm_Initialize()
    ' Purpose: Dynamically create the Label, TextBox, and Button, then load data.
    
    ' --- Layout Constants ---
    Const CONTROL_WIDTH As Single = 220
    Const INPUT_HEIGHT As Single = 18
    Const BTN_HEIGHT As Single = 24
    Const PADDING As Single = 10
    Const LBL_HEIGHT As Single = 12
    
    Dim currentTop As Single
    currentTop = PADDING
    
    ' --- 1. Create Label ---
    Set lblEndpoint = Me.Controls.Add("Forms.Label.1", "lblEndpoint")
    With lblEndpoint
        .Caption = "API Endpoint URL:"
        .Left = PADDING
        .Top = currentTop
        .Width = CONTROL_WIDTH
        .Height = LBL_HEIGHT
    End With
    
    currentTop = currentTop + LBL_HEIGHT + 2 ' Small gap
    
    ' --- 2. Create TextBox ---
    Set txtApiEndpoint = Me.Controls.Add("Forms.TextBox.1", "txtApiEndpoint")
    With txtApiEndpoint
        .Left = PADDING
        .Top = currentTop
        .Width = CONTROL_WIDTH
        .Height = INPUT_HEIGHT
        ' Add a border for better visibility
        .BorderStyle = fmBorderStyleSingle
    End With
    
    currentTop = currentTop + INPUT_HEIGHT + PADDING
    
    ' --- 3. Create Save Button ---
    Set btnSave = Me.Controls.Add("Forms.CommandButton.1", "btnSave")
    With btnSave
        .Caption = "Save Settings"
        .Left = PADDING
        .Top = currentTop
        .Width = 80
        .Height = BTN_HEIGHT
        .Accelerator = "S" ' Allows Alt+S to trigger
    End With
    
    ' --- 4. Final Form Sizing ---
    ' Adjust form size to fit controls perfectly
    Me.Width = CONTROL_WIDTH + (PADDING * 2) + 12 ' +12 accounts for window borders
    Me.Height = currentTop + BTN_HEIGHT + 30 + 10 ' +30 accounts for title bar, +10 padding to bottom
    Me.Caption = "LnAddinPro Settings"
    
    ' --- 5. Load Data ---
    LoadSettingsData
End Sub

Private Sub LoadSettingsData()
    ' Purpose: Read the JSON file and populate the TextBox
    
    ' Get config path via the Manager Module
    m_ConfigPath = mod_config_mgr.GetConfigPath()
    
    ' Load JSON into memory
    Set m_JsonData = mod_config_mgr.LoadConfigFromJson(m_ConfigPath)
    
    ' Safely extract the data
    ' Path: root -> "config" -> "api_endpoint"
    If Not m_JsonData Is Nothing Then
        If m_JsonData.Exists("config") Then
            If m_JsonData("config").Exists("api_endpoint") Then
                txtApiEndpoint.Text = m_JsonData("config")("api_endpoint")
            End If
        End If
    Else
        ' Fallback default if file is missing or corrupt
        txtApiEndpoint.Text = "https://"
    End If
End Sub

Private Sub btnSave_Click()
    ' Purpose: Save the input back to the JSON file
    
    Dim newUrl As String
    newUrl = Trim(txtApiEndpoint.Text)
    
    ' --- Validation ---
    If Len(newUrl) = 0 Then
        MsgBox "API Endpoint cannot be empty.", vbExclamation, "Validation Error"
        txtApiEndpoint.SetFocus
        Exit Sub
    End If
    
    ' --- Update Memory Object ---
    ' Ensure the Dictionary structure exists
    If m_JsonData Is Nothing Then Set m_JsonData = New Scripting.Dictionary
    
    If Not m_JsonData.Exists("config") Then
        m_JsonData.Add "config", New Scripting.Dictionary
    End If
    
    ' Update the key
    m_JsonData("config")("api_endpoint") = newUrl
    
    ' --- Write to File ---
    ' Call the helper function in your standard module
    mod_config_mgr.SaveConfigToJson m_ConfigPath, m_JsonData
    
    MsgBox "Configuration saved successfully.", vbInformation, "LnAddinPro"
    
    ' Close the form
    Unload Me
End Sub