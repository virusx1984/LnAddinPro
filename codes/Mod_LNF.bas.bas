Attribute VB_Name = "Mod_LNF"
' This function converts a given Excel range into a Markdown table string.
' It's useful for quickly formatting data for documentation or code repositories.
'
' @param sourceRange (Range): The input range to be converted. This is a reference to the actual range.
' @param headerRowsCount (Long): The number of header rows. A separator line will be inserted after this row. Defaults to 1.
' @param columnDelimiter (String): The character to separate columns. Defaults to "|".
' @param rowDelimiter (String): The character to separate rows. Defaults to a newline.
' @return (String): A string formatted as a Markdown table.
Public Function LNF_CreateTable(ByVal sourceRange As Range, _
                                Optional ByVal headerRowsCount As Long = 1, _
                                Optional ByVal columnDelimiter As String = "|", _
                                Optional ByVal rowDelimiter As String = "") As String

    ' Check if delimiters are empty and set to defaults if necessary.
    ' Using ByVal for parameters allows for this without affecting the original variable.
    If columnDelimiter = "" Then columnDelimiter = "|"
    If rowDelimiter = "" Then rowDelimiter = Chr(10)

    ' Use descriptive variable names and declare their types for clarity and performance.
    Dim totalRows As Long
    Dim totalColumns As Long
    Dim markdownTableString As String
    Dim currentRowString As String
    Dim colIndex As Long
    Dim rowIndex As Long

    totalRows = sourceRange.Rows.Count
    totalColumns = sourceRange.Columns.Count

    ' Build the Markdown table string row by row.
    For rowIndex = 1 To totalRows
        currentRowString = ""

        ' Loop through each column in the current row.
        For colIndex = 1 To totalColumns
            ' Get the text from the cell and add the column delimiter.
            ' Ensure the string starts and ends with a delimiter for valid Markdown.
            currentRowString = currentRowString & columnDelimiter & sourceRange.Cells(rowIndex, colIndex).Text
        Next colIndex
        
        ' Add closing delimiter and row delimiter.
        currentRowString = currentRowString & columnDelimiter & rowDelimiter
        
        ' Append the formatted row to the main result string.
        markdownTableString = markdownTableString & currentRowString
        
        ' Check if a header separator line needs to be inserted.
        If rowIndex = headerRowsCount Then
            Dim separatorRowString As String
            separatorRowString = ""
            
            For colIndex = 1 To totalColumns
                separatorRowString = separatorRowString & columnDelimiter & "---"
            Next colIndex

            ' Add closing delimiter and row delimiter for the separator line.
            separatorRowString = separatorRowString & columnDelimiter & rowDelimiter
            
            markdownTableString = markdownTableString & separatorRowString
        End If
    Next rowIndex
    
    ' Return the final Markdown table string.
    LNF_CreateTable = markdownTableString
    
End Function



' This function converts a given Excel range, including merged cells, into an HTML table string.
' It's highly useful for converting Excel data into web-friendly formats.
'
' @param sourceRange (Range): The input range to be converted.
' @param headerRowsCount (Long): The number of rows to be treated as header cells (<th>).
' @return (String): A string formatted as an HTML table.
Public Function LNF_RangeToHTML(ByVal sourceRange As Range, ByVal headerRowsCount As Long) As String
    
    Dim html As String
    Dim currentRow As Long, currentColumn As Long
    Dim cell As Range
    
    ' Use a more modern and robust way to handle dictionary objects.
    ' This requires adding a reference to "Microsoft Scripting Runtime".
    ' Go to Tools -> References in the VBA editor, and check the box for "Microsoft Scripting Runtime".
    Dim processedCells As Scripting.Dictionary
    Set processedCells = New Scripting.Dictionary
        
    html = "<table border='1'>" & vbCrLf
    
    For currentRow = 1 To sourceRange.Rows.Count
        html = html & "<tr>" & vbCrLf
        
        For currentColumn = 1 To sourceRange.Columns.Count
            Set cell = sourceRange.Cells(currentRow, currentColumn)
            
            ' Skip cells that are part of a merged area already processed.
            If Not processedCells.Exists(cell.Address(External:=True)) Then
                
                Dim rowspanCount As Long
                Dim colspanCount As Long
                
                ' Get the dimensions of the merged area.
                ' Use a temporary range variable to prevent errors if the cell is not merged.
                Dim mergedArea As Range
                Set mergedArea = cell.MergeArea
                rowspanCount = mergedArea.Rows.Count
                colspanCount = mergedArea.Columns.Count
                
                ' Determine the HTML tag based on whether it's a header row or a data row.
                Dim cellTag As String
                If currentRow <= headerRowsCount Then
                    cellTag = "th"
                Else
                    cellTag = "td"
                End If
                
                ' Build the HTML for the current cell.
                html = html & "<" & cellTag
                
                If rowspanCount > 1 Then html = html & " rowspan='" & rowspanCount & "'"
                If colspanCount > 1 Then html = html & " colspan='" & colspanCount & "'"
                
                html = html & ">" & cell.Text & "</" & cellTag & ">" & vbCrLf
                
                ' Mark all cells in the merged area as processed to prevent re-processing.
                Dim mergedCell As Range
                For Each mergedCell In mergedArea
                    ' Using the full address including sheet name makes the key more robust.
                    processedCells.Add mergedCell.Address(External:=True), True
                Next mergedCell
            End If
        Next currentColumn
        
        html = html & "</tr>" & vbCrLf
    Next currentRow
    
    html = html & "</table>"
    LNF_RangeToHTML = html
End Function


' This function concatenates the text from a given range of cells.
' It allows for a custom join string and optional surrounding characters for each cell's text.
'
' @param sourceRange (Range): The range of cells whose text is to be joined.
' @param joinString (String): The string used to separate the text from each cell.
' @param leftSurround (Optional String): The character(s) to place at the beginning of each cell's text.
' @param rightSurround (Optional String): The character(s) to place at the end of each cell's text.
' @return (String): The concatenated string.
Public Function LNF_Join(ByVal sourceRange As Range, ByVal joinString As String, _
                         Optional ByVal leftSurround As String = "", _
                         Optional ByVal rightSurround As String = "") As String
    
    Dim cell As Range
    Dim resultString As String
    
    ' If the right surround string is not provided, use the left surround string.
    If rightSurround = "" Then rightSurround = leftSurround
    
    ' Loop through each cell in the provided range.
    For Each cell In sourceRange
        ' Avoid processing empty cells unless specifically desired.
        ' This prevents extra delimiters in the final string.
        If Not IsEmpty(cell.Value) And Len(cell.Value) > 0 Then
            ' Append the join string if the result string is not empty.
            If Len(resultString) > 0 Then
                resultString = resultString & joinString
            End If
            
            ' Append the surrounding characters and the cell's text.
            resultString = resultString & leftSurround & CStr(cell.Value) & rightSurround
        End If
    Next cell
    
    LNF_Join = resultString
    
End Function
