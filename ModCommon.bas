Attribute VB_Name = "ModCommon"
Option Explicit


'=================================================
'check whether is the same pattern by regex.
'=================================================
Function IsTestOK(inputString As String, pattern As String) As Boolean
    Dim reg As New RegExp
    reg.pattern = pattern
    reg.IgnoreCase = True
    IsTestOK = reg.Test(inputString)
End Function

Public Function GetMatchCollection(strInput As String, pattern As String) As MatchCollection
    Dim reg As New RegExp
    
    reg.IgnoreCase = True
    reg.pattern = pattern
    reg.Global = True
    
    Dim ms As MatchCollection
    Set ms = reg.Execute(strInput)
    Set GetMatchCollection = ms
    
End Function

'=================================================
'Get match value by pattern using regex
'=================================================
Function GetMatchValue(inputString As String, pattern As String, Optional groupID As Integer = 0) As String
    Dim reg As New RegExp
    reg.pattern = pattern
    reg.IgnoreCase = True
    
    Dim ms As MatchCollection
    Set ms = reg.Execute(inputString)
    If ms.Count > 0 Then
        Dim m As Match
        Set m = ms(0)
        If groupID = 0 Then
            GetMatchValue = m.value
        Else
            GetMatchValue = m.SubMatches(groupID - 1)
        End If
    End If
End Function

'=================================================
'Replace string by pattern. replace all matches
'=================================================
Function ReplaceMatch(inputString As String, pattern As String, replaceString As String) As String
    Dim reg As New RegExp
    reg.pattern = pattern
    reg.IgnoreCase = True
    reg.MultiLine = False
    reg.Global = True
    ReplaceMatch = reg.Replace(inputString, replaceString)
    
End Function

'=================================================
'Replace string by pattern. replace all matches
'=================================================
Function ReplaceMatchWithoutCR(inputString As String, pattern As String, replaceString As String) As String
    Dim reg As New RegExp
    reg.IgnoreCase = True
    reg.MultiLine = False
    reg.Global = True
    
    reg.pattern = vbLf
    Dim temp As String
    temp = reg.Replace(replaceString, "")
    
    reg.pattern = pattern
    ReplaceMatchWithoutCR = reg.Replace(inputString, temp)
    
End Function


'=================================================
'Get the last word from a string
'=================================================
Function LastWord(inputString As String) As String
    LastWord = GetMatchValue(inputString, "\w+$")
End Function

'=================================================
'Get File name from file full Path
'=================================================
Function GetFileName(strPath As String) As String
    GetFileName = GetMatchValue(strPath, "[^\\\/]+$")
End Function

'=================================================
'Clear All items in collection
'=================================================
Sub RemoveAll(list As Collection)
    If list Is Nothing Then Exit Sub
    
    While list.Count > 0
        list.Remove 1
    Wend

End Sub

'=================================================
'Set Hyperlink. delete when exists, then add.
'=================================================
Sub ResetHyperLink(thisRange As Range, backRange As Range)
    
    If thisRange.Hyperlinks.Count > 0 Then
        thisRange.Hyperlinks.Delete
    End If
    
    Dim thisSHeet As Worksheet
    Set thisSHeet = thisRange.Worksheet
    
    Dim toWhere As String
    toWhere = "'" & backRange.Worksheet.name & "'!" & backRange.Address
    
    'anchor: From Where, Address="" is thisworkbook inner.
    'SubAddress:=sheetname!address. if sheetname emited means selfsheet.
    thisSHeet.Hyperlinks.Add anchor:=thisRange, Address:="", SubAddress:=toWhere, TextToDisplay:=thisRange.Text

End Sub

Public Sub ReadCSVToSheet(thisSHeet As Worksheet, csvPath As String, Optional firstCell As String = "A1")
    thisSHeet.Activate
    Dim maxRow As Integer
    Dim maxCol As Integer
    maxRow = thisSHeet.UsedRange.Rows.Count
    maxCol = thisSHeet.UsedRange.Columns.Count
    thisSHeet.Range(thisSHeet.Range(firstCell), thisSHeet.Cells(maxRow, maxCol)).ClearContents
    
    '外部データの取り込みでデータを取得
    With thisSHeet.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=thisSHeet.Range(firstCell))
        .FieldNames = False
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = False
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .refresh BackgroundQuery:=False
        .Delete
    End With
    
End Sub

Public Sub SetTableStyle(thisSHeet As Worksheet, TableName As String)
    thisSHeet.Activate
    
    If IsEmpty(thisSHeet.UsedRange.Value2) Then
        Exit Sub
    End If
    
    thisSHeet.ListObjects.Add(xlSrcRange, thisSHeet.UsedRange, , xlYes).name _
        = TableName
    thisSHeet.ListObjects(TableName).TableStyle = "TableStyleLight21"

    thisSHeet.Range(TableName & "[#Headers]").Select
    
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub

Public Function ReplaceInvaidWord(words As String) As String
    Dim tName As String
    tName = Replace(words, ".", "")
    tName = Replace(tName, "-", "")
    ReplaceInvaidWord = tName
End Function

Public Function GetRangeValues(thisSHeet As Worksheet, rangeName As String) As Variant
    Dim cRange As Range
    Set cRange = FindInSheet(thisSHeet, rangeName)
    GetRangeValues = cRange.CurrentRegion.value
End Function


Public Function GetSheetName(thisRange As Range) As String
    GetSheetName = thisRange.Worksheet.name
End Function

Public Sub MkDirAll(strPath As String)
    Dim subPath() As String
    strPath = Replace(strPath, "/", "\")
    subPath = Split(strPath, "\")
    Dim p
    Dim rootPath As String
    For Each p In subPath
        rootPath = rootPath & p & "\"
        If Not fso.FolderExists(rootPath) Then
            MkDir rootPath
        End If
    Next

End Sub
