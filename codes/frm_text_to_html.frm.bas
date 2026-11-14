Option Explicit

' --- Event Handlers for Dynamic Controls ---
Public WithEvents btnExport As MSForms.CommandButton
Public WithEvents btnClose As MSForms.CommandButton

' --- Control References ---
Private txtContent As MSForms.TextBox
Private lblStatus As MSForms.Label

' ==============================================================================
' UI INITIALIZATION
' ==============================================================================
Private Sub UserForm_Initialize()
    Me.Caption = "Text/Code to HTML Exporter"
    Me.Width = 500
    Me.Height = 400
    
    Dim currentTop As Long: currentTop = 10
    Const MARGIN As Long = 10
    
    ' 1. Label
    With Me.Controls.Add("Forms.Label.1", "lblDesc")
        .Caption = "Paste your Chinese text or code below:"
        .Left = MARGIN
        .Top = currentTop
        .Width = Me.InsideWidth - (MARGIN * 2)
    End With
    currentTop = currentTop + 15
    
    ' 2. TextBox (MultiLine)
    Set txtContent = Me.Controls.Add("Forms.TextBox.1", "txtContent")
    With txtContent
        .Left = MARGIN
        .Top = currentTop
        .Width = Me.InsideWidth - (MARGIN * 2)
        .Height = Me.InsideHeight - currentTop - 60
        .MultiLine = True
        .ScrollBars = fmScrollBarsBoth
        .WordWrap = False
        .Font.Name = "Consolas"
        .Font.Size = 10
    End With
    currentTop = txtContent.Top + txtContent.Height + 10
    
    ' 3. Buttons
    Set btnExport = Me.Controls.Add("Forms.CommandButton.1", "btnExport")
    With btnExport
        .Caption = "Export to HTML"
        .Left = MARGIN
        .Top = currentTop
        .Width = 120
        .Height = 28
        .BackColor = &H80FF80
        .Font.bold = True
    End With
    
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close"
        .Left = MARGIN + 130
        .Top = currentTop
        .Width = 80
        .Height = 28
    End With
    
    ' 4. Status Label
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus")
    With lblStatus
        .Caption = "Ready."
        .Left = MARGIN + 220
        .Top = currentTop + 6
        .Width = Me.InsideWidth - (MARGIN * 2) - 220
        .ForeColor = &H808080
    End With
End Sub

' ==============================================================================
' ACTIONS
' ==============================================================================
Private Sub btnExport_Click()
    Dim rawText As String
    Dim safeHtml As String
    Dim htmlContent As String
    Dim filePath As Variant
    
    rawText = txtContent.Text
    If Len(rawText) = 0 Then
        MsgBox "Please paste some text or code first.", vbExclamation
        Exit Sub
    End If
    
    ' 1. HTML Entity Escaping (Order matters: & first)
    safeHtml = rawText
    safeHtml = Replace(safeHtml, "&", "&amp;")
    safeHtml = Replace(safeHtml, "<", "&lt;")
    safeHtml = Replace(safeHtml, ">", "&gt;")
    
    ' 2. Build HTML Template
    htmlContent = "<!DOCTYPE html>" & vbCrLf & _
                  "<html>" & vbCrLf & _
                  "<head>" & vbCrLf & _
                  "    <meta charset=""UTF-8"">" & vbCrLf & _
                  "    <title>Exported Text/Code</title>" & vbCrLf & _
                  "    <style>" & vbCrLf & _
                  "        body { font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; background: #f9f9f9; }" & vbCrLf & _
                  "        pre { background: #ffffff; border: 1px solid #ccc; padding: 15px; border-radius: 5px; overflow-x: auto; font-family: Consolas, 'Courier New', monospace; font-size: 14px; }" & vbCrLf & _
                  "    </style>" & vbCrLf & _
                  "</head>" & vbCrLf & _
                  "<body>" & vbCrLf & _
                  "    <pre>" & safeHtml & "</pre>" & vbCrLf & _
                  "</body>" & vbCrLf & _
                  "</html>"
                  
    ' 3. Ask for save location
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="ExportedText.html", _
        FileFilter:="HTML Files (*.html), *.html", _
        Title:="Save HTML File")
        
    If filePath = False Then Exit Sub
    
    ' 4. Write to file using ADODB.Stream (UTF-8)
    On Error GoTo ErrHandler
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2             ' adTypeText
        .Charset = "UTF-8"    ' Ensure Chinese characters don't garble
        .Open
        .WriteText htmlContent
        .SaveToFile CStr(filePath), 2 ' 2 = adSaveCreateOverWrite
        .Close
    End With
    Set stream = Nothing
    
    lblStatus.Caption = "Exported successfully at " & Format(Now, "hh:mm:ss")
    lblStatus.ForeColor = &H8000& ' Green
    
    ' Ask if user wants to open it immediately
    If MsgBox("Export successful! Do you want to open the HTML file now?", vbQuestion + vbYesNo) = vbYes Then
        Shell "explorer.exe """ & filePath & """", vbNormalFocus
    End If
    
    Exit Sub

ErrHandler:
    lblStatus.Caption = "Error: " & Err.Description
    lblStatus.ForeColor = &HFF& ' Red
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub