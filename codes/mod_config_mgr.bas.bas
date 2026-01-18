Attribute VB_Name = "mod_config_mgr"
' Module: mod_config_mgr
Option Explicit

Public Const CONFIG_FILE_NAME As String = "LnAddinPro.json"

' Main entry point for the Ribbon button
Public Sub ShowSettingsDialog()
    Dim configPath As String
    configPath = GetConfigPath()
    
    ' Check if config file exists; create default if missing
    If Not FileExists(configPath) Then
        CreateDefaultConfig configPath
    End If
    
    ' Open the settings form
    frm_settings.Show
End Sub

' Helper: Get the full file path for the config file
Public Function GetConfigPath() As String
    GetConfigPath = ThisWorkbook.Path & Application.PathSeparator & CONFIG_FILE_NAME
End Function

' Helper: Check if a file exists
Public Function FileExists(filePath As String) As Boolean
    Dim fso As New Scripting.FileSystemObject
    FileExists = fso.FileExists(filePath)
End Function

' Create the default JSON structure
Private Sub CreateDefaultConfig(filePath As String)
    Dim root As New Scripting.Dictionary
    Dim config As New Scripting.Dictionary
    
    ' Build structure: {"config": {"api_endpoint": "https://"}}
    config.Add "api_endpoint", "https://"
    root.Add "config", config
    
    ' Save to file
    SaveConfigToJson filePath, root
End Sub

' Read JSON file and return a Dictionary object
Public Function LoadConfigFromJson(filePath As String) As Scripting.Dictionary
    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim jsonText As String
    
    If Not fso.FileExists(filePath) Then Exit Function
    
    ' Open file for reading
    Set ts = fso.OpenTextFile(filePath, ForReading)
    jsonText = ts.ReadAll
    ts.Close
    
    ' Parse using VBA-JSON library
    Set LoadConfigFromJson = JsonConverter.ParseJson(jsonText)
End Function

' Save a Dictionary object to a JSON file
Public Sub SaveConfigToJson(filePath As String, jsonDict As Scripting.Dictionary)
    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim jsonString As String
    
    ' Convert Dictionary to JSON string (Whitespace:=2 adds indentation)
    jsonString = JsonConverter.ConvertToJson(jsonDict, Whitespace:=2)
    
    ' Write to file (True = Overwrite)
    Set ts = fso.CreateTextFile(filePath, True)
    ts.Write jsonString
    ts.Close
End Sub
