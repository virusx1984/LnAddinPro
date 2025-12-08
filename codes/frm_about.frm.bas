Attribute VB_Name = "frm_about"
Attribute VB_Base = "0{9C33C3BD-02D4-4C8C-8476-85BCE954677D}{E611F722-C076-40F5-AA35-02990847B0A3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' UserForm: frm_about
' Purpose: Display Add-in information (Version, Author, Copyright).
Option Explicit

Public WithEvents btnOK As MSForms.CommandButton
Attribute btnOK.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    ' --- Setup Form ---
    Me.Caption = "About LnAddinPro"
    Me.Width = 300
    Me.Height = 220
    Me.BackColor = &HFFFFFF ' White background for clean look
    
    Dim currentTop As Long
    Const MARGIN As Long = 20
    Const LABEL_W As Long = 240
    
    currentTop = MARGIN
    
    ' --- 1. Title (Large, Bold) ---
    With Me.Controls.Add("Forms.Label.1", "lblTitle")
        .Caption = "LnAddinPro"
        .Left = MARGIN: .Top = currentTop: .Width = LABEL_W: .Height = 30
        .Font.Name = "Segoe UI"
        .Font.Size = 18
        .Font.Bold = True
        .ForeColor = &H217346 ' Excel Green
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
    End With
    currentTop = currentTop + 35
    
    ' --- 2. Version ---
    With Me.Controls.Add("Forms.Label.1", "lblVersion")
        .Caption = "Version 1.0.2"
        .Left = MARGIN: .Top = currentTop: .Width = LABEL_W: .Height = 15
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .ForeColor = &H808080 ' Grey
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
    End With
    currentTop = currentTop + 25
    
    ' --- 3. Horizontal Line (Visual Separator) ---
    With Me.Controls.Add("Forms.Label.1", "lblLine")
        .Caption = ""
        .Left = MARGIN + 20: .Top = currentTop: .Width = LABEL_W - 40: .Height = 1
        .BackColor = &HE0E0E0
    End With
    currentTop = currentTop + 15
    
    ' --- 4. Author & Description ---
    With Me.Controls.Add("Forms.Label.1", "lblDesc")
        .Caption = "A powerful Excel toolkit for data reshaping, analysis, and code generation." & vbCrLf & vbCrLf & _
                   "Author: Lining (virusx1984@gmail.com)" & vbCrLf & _
                   "Copyright ? 2025"
        .Left = MARGIN: .Top = currentTop: .Width = LABEL_W: .Height = 60
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
    End With
    currentTop = currentTop + 60
    
    ' --- 5. OK Button ---
    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    With btnOK
        .Caption = "OK"
        .Left = (Me.InsideWidth - 80) / 2 ' Center button
        .Top = currentTop
        .Width = 80: .Height = 25
        .BackColor = &HF0F0F0
    End With
End Sub

Private Sub btnOK_Click()
    Unload Me
End Sub

