Attribute VB_Name = "ModBook"
Option Explicit

'================================================
' Workbook SaveAs
'================================================
Function BookSaveAs(thisBook As Workbook, strFileName As String) As Integer
On Error GoTo err1:

    thisBook.SaveAs strFileName, CreateBackup:=False
    Exit Function
    
err1:
    BookSaveAs = -1
End Function
'================================================
' Open Workbook
'================================================
Function OpenBook(strFileName As String, Optional AddifNotFound As Boolean = True) As Workbook
    Dim book As Workbook
    
    For Each book In Workbooks
        If book.name = strFileName Then
            Set OpenBook = book
            Exit Function
        End If
        If book.FullName = strFileName Then
            Set OpenBook = book
            Exit Function
        End If
    Next
    
    If fso.FileExists(strFileName) Then
        Application.DisplayAlerts = False
        Set OpenBook = Workbooks.Open(strFileName, readonly:=False, IgnoreReadOnlyRecommended:=True, Editable:=True)
    ElseIf Not AddifNotFound Then
        '開いたWorkbookだけを探す場合、見つからなければ、終了。
        Exit Function
    Else
        Set book = Workbooks.Add()
        BookSaveAs book, strFileName
        Set OpenBook = book
    End If
    
End Function
'================================================
' Workbook SaveAs
'================================================
Sub BookCloseAndSave(thisBook As Workbook, Optional fullpath As String = "")
Dim showAlert As Boolean
    
    showAlert = Application.DisplayAlerts
    
    Application.DisplayAlerts = False
    On Error Resume Next
    If fullpath = "" Then
        thisBook.Close True
    Else
        BookSaveAs thisBook, fullpath
        thisBook.Close
    End If
    Application.DisplayAlerts = showAlert
    
End Sub

'================================================
' Workbook Close by Name
'================================================
Sub BookCloseAndSaveByName(workbookName As String)
    Dim book As Workbook
    For Each book In Workbooks
        If book.name = workbookName Then
            BookCloseAndSave book
            Exit Sub
        End If
        If book.FullName = workbookName Then
            BookCloseAndSave book
            Exit Sub
        End If
    Next
    
End Sub

'================================================
' Workbook Close by Names
'================================================
Sub BookCloseAndSaveByNames(workbookNames() As String)
    Dim i%
    For i = 0 To UBound(workbookNames)
       BookCloseAndSaveByName workbookNames(i)
    Next
    
End Sub

'================================================
' Close Workbook Without Save and no alert.
'================================================
Function BookCloseNoSave(book As Workbook) As Boolean
    
    Dim showAlert As Boolean
    
    showAlert = Application.DisplayAlerts
    
    Application.DisplayAlerts = False
    book.Close False
    Application.DisplayAlerts = showAlert
    
End Function

Public Function HasSheet(SheetName As String, Optional book As Workbook = Nothing) As Boolean
    If book Is Nothing Then
        Set book = ThisWorkbook
    End If
    
    Dim sheet As Worksheet
    For Each sheet In book.Sheets
        If sheet.name = SheetName Then
            HasSheet = True
            Exit Function
        End If
    Next
End Function

Public Function GetSheet(SheetName As String, Optional book As Workbook = Nothing, Optional AddifNotFound = True) As Worksheet
    If book Is Nothing Then
        Set book = ThisWorkbook
    End If
    
    If HasSheet(SheetName, book) Then
        Set GetSheet = book.Sheets(SheetName)
        Exit Function
    End If
    If AddifNotFound Then
        Dim sheet As Worksheet
        Set sheet = book.Sheets.Add
        sheet.name = SheetName
        Set GetSheet = sheet
    End If
End Function
 
Public Sub CopyData(fromSheet As Worksheet, toSheet As Worksheet, r As Integer, Optional title As String = "")
    Dim rx, cx As Integer
    rx = fromSheet.UsedRange.Rows.Count
    cx = fromSheet.UsedRange.Columns.Count
    Dim fromRange As Range
    Set fromRange = fromSheet.Range(fromSheet.Cells(1, 1), fromSheet.Cells(rx, cx))
    fromRange.Copy
    
    toSheet.Activate
    toSheet.Range("B" & r).Select
    toSheet.Paste
    toSheet.Range("A" & r) = title
    If 1 < rx Then
        toSheet.Range("A" & r & ":A" & (r + rx - 1)).FillDown
    End If
End Sub

Public Sub ShowSheetTable(thisSHeet As Worksheet)
    
    Dim r, c As Integer
    r = thisSHeet.UsedRange.Rows.Count
    c = thisSHeet.UsedRange.Columns.Count
    Dim thisRange As Range
    Set thisRange = thisSHeet.Range(thisSHeet.Cells(1, 1), thisSHeet.Cells(r, c))
    
    thisSHeet.ListObjects.Add(xlSrcRange, thisRange, , xlYes).name = thisSHeet.name
    thisSHeet.ListObjects(thisSHeet.name).TableStyle = "TableStyleLight21"
    
End Sub


Public Sub ExportReport(thisSHeet As Worksheet, reportFolder As String)
    'Save To XlsReport
    If Not fso.FolderExists(reportFolder) Then
        fso.CreateFolder reportFolder
    End If
    
    SaveAsNewXls thisSHeet, fso.BuildPath(reportFolder, thisSHeet.name & Format(Now, "YYYYMMDDHHMMSS") & ".xlsx")
    
End Sub

Sub SaveAsNewXls(thisSHeet As Worksheet, REPORTPATH As String)
    
    thisSHeet.Activate
    thisSHeet.Copy
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs REPORTPATH
    ActiveWindow.Close
    Application.DisplayAlerts = True
End Sub


Public Function findrange(thisRange As Range, findKey As String, Optional afterCell As Range = Nothing, Optional lookAt As Integer = xlPart) As Range
    Dim thisSHeet As Worksheet
    Set thisSHeet = thisRange.Worksheet
    If afterCell Is Nothing Then
        Set afterCell = thisSHeet.Cells(thisRange.Row, thisRange.Column)
    End If
    
    Set findrange = thisRange.Find(What:=findKey, After:=afterCell, lookAt:=lookAt)
End Function

Public Function FindInSheet(thisSHeet As Worksheet, findKey As String, Optional afterCell As Range = Nothing) As Range
    Dim thisRange As Range
    If afterCell Is Nothing Then
        Set afterCell = thisSHeet.Cells(1, 1)
    End If
    
    Set FindInSheet = thisSHeet.Cells.Find(What:=findKey, After:=afterCell, lookAt:=xlPart)
End Function

Public Function GetFilePath(para As String) As String
    Dim filePath$
    filePath = ThisWorkbook.Path
    If (para <> "") Then
        filePath = filePath + "\" + para
    End If
    GetFilePath = filePath
End Function
