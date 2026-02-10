Sub TestJsonWithFormat()
    Dim rng As Range
    Dim typeDict As Object
    Dim fmtDict As Object
    Dim jsonStr As String
    
    Set rng = Range("R15:X20")
    
    ' 1. Types
    Set typeDict = CreateObject("Scripting.Dictionary")
    typeDict.Add "col1", "cate"
    typeDict.Add "col2", "cate"
    typeDict.Add "col3", "cate"
    typeDict.Add "col4", "cate"
    
    ' ... others ...
    
    ' 2. Formats (New)
    Set fmtDict = CreateObject("Scripting.Dictionary")
    
    ' Example 1: Force fixed decimals (Will output as JSON Number: 1200.00)
    fmtDict.Add "col5", "0.00"
    
    ' Example 2: With separator (Will output as JSON String: "1,500")
    fmtDict.Add "col6", "#,##0"
    
    ' Example 3: Percent (Will output as JSON String: "50%")
    fmtDict.Add "col7", "0%"

    jsonStr = RangeToCompactJson(rng, typeDict, fmtDict)
    
    Debug.Print jsonStr
End Sub