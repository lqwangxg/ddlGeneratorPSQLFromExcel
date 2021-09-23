Attribute VB_Name = "generateDDL"
Option Explicit
Public fso As New FileSystemObject


'ignoreCase_: ingnore upper or lower cases, global_: one pattern string is matched multiple times

Private Function getRegexp(target, matchPattern, Optional ignoreCase_ = True, Optional global_ = True)
    
    Dim regex As RegExp
    Dim matches As MatchCollection
    
    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .pattern = matchPattern
        .IgnoreCase = ignoreCase_
        .Global = global_
        Set matches = .Execute(target)
    End With

    If matches.Count <> 0 Then
        getRegexp = matches(0)
    Else
        getRegexp = ""
    End If
    
    Set regex = Nothing
    Set matches = Nothing

End Function


Private Function CreateTable(saveName, tableHeader As tableHeader)
    Dim Str As String
    Str = ""
    Dim TableName As String
    TableName = Range(tableHeader.cellTableName).value
    Dim fields As String
    fields = ""
    Dim alters As String
    alters = ""
    Dim lineNo As Integer: lineNo = tableHeader.lineNoFirstCol
    Dim pkey: pkey = ""
    Do
        Dim nn As String
        If StrComp("y", Range(tableHeader.rowNotNull & lineNo).value) = 0 _
          Or Range(tableHeader.rowNotNull & lineNo).value = "Åõ" Then
            nn = ""
        Else
            nn = " NOT NULL"
        End If
        
        Dim dtype As String
        Dim tVal As String
        tVal = Range(tableHeader.rowDType & lineNo).value
        If StrComp("varchar", tVal) = 0 Then
            Dim dlen As String: dlen = Range(tableHeader.rowLen & lineNo).value
            If dlen = "" Then
                MsgBox "length n of varchar(n) is not specified."
                Exit Function
            End If
            dtype = "character varying(" & dlen & ")"
        ElseIf StrComp("char", tVal) = 0 Then
            dtype = tVal
        ElseIf StrComp("serial", tVal) = 0 Then
            dtype = tVal
        ElseIf StrComp("boolean", tVal) = 0 Then
            dtype = tVal
        ElseIf StrComp("integer", tVal) = 0 Then
            dtype = "integer"
        ElseIf StrComp("timestamp", tVal) = 0 Then
            dtype = "timestamp with time zone"
        ElseIf StrComp("smallint", tVal) = 0 Then
            dtype = tVal
        ElseIf StrComp("time", tVal) = 0 Then
            dtype = "time with time zone"
        ElseIf StrComp("date", tVal) = 0 Then
            dtype = tVal
        ElseIf StrComp("text", tVal) = 0 Then
            dtype = tVal
        ElseIf StrComp("bytea", tVal) = 0 Then
            dtype = tVal
        Else
            MsgBox "Unknown Data Type:" & tVal & " on " & Range(tableHeader.rowDType & lineNo).Address
            End
        End If
        
        If Len(fields) <> 0 Then
            fields = fields & ","
        End If
        
        Dim ColumnName As String: ColumnName = Range(tableHeader.rowColName & lineNo).value
        fields = fields & " " & ColumnName & " " & dtype & nn & vbNewLine
        
        ' Primary Key
        If IsTestOK(Range(tableHeader.rowPkey & lineNo).value, "P") Then
            If Len(pkey) <> 0 Then
                pkey = pkey & ","
            End If
            pkey = pkey & ColumnName
        End If
    
        Dim fkWork: fkWork = Range(tableHeader.rowConstr & lineNo).value
        
        ' Unique
        If InStr(fkWork, "UNIQUE") <> 0 Then
        
            Dim unique: unique = ""
            
            unique = getRegexp(fkWork, "UNIQUE\(.*\)", False)

            If unique <> "" Then
                unique = Replace(fkWork, "UNIQUE", "")
                alters = alters & "ALTER TABLE ONLY " & TableName & " ADD CONSTRAINT m_" & TableName & "_" & ColumnName & "_uq UNIQUE " & unique & ";" & vbNewLine
            Else
                alters = alters & "ALTER TABLE ONLY " & TableName & " ADD CONSTRAINT m_" & TableName & "_" & ColumnName & "_uq UNIQUE (" & ColumnName & ");" & vbNewLine
            End If
        End If
        
        ' References
        If InStr(fkWork, "REFERENCES") <> 0 Then
        
            Dim references: references = ""
            references = getRegexp(fkWork, "REFERENCES\(.*\)", False)

            If references <> "" Then
                Dim tblName: tblName = Replace(references, "REFERENCES(", "")
                tblName = Replace(tblName, ")", "")
                Dim colName: colName = "id"

                'Set Foreign Key
                alters = alters & "ALTER TABLE ONLY " & TableName & " ADD CONSTRAINT fk_" & TableName & "_" & ColumnName & " FOREIGN KEY (" & ColumnName & ") REFERENCES " & tblName & "(" & colName & ");" & vbNewLine

            End If
        End If
    
        ' Comment on each column
        alters = alters & "COMMENT ON COLUMN " & TableName & "." & ColumnName & " IS '" & Range(tableHeader.rowCommentCol & lineNo).value & "';" & vbNewLine

        lineNo = lineNo + 1
        
    Loop While Range(tableHeader.rowColName & lineNo).value <> ""
    
    ' Comment on table
    If Len(pkey) <> 0 Then
        alters = alters & "ALTER TABLE ONLY " & TableName & " ADD CONSTRAINT m_" & TableName & "_pkey PRIMARY KEY (" & pkey & ");" & vbNewLine
    End If
    alters = alters & "COMMENT ON TABLE " & TableName & " IS '" & Range(tableHeader.rowCommentTbl).value & "';" & vbNewLine
    'alters = alters & "ALTER TABLE public." & TableName & " OWNER TO " & tableHeader.ownerName & ";" & vbNewLine
    
    '
    Str = Str & "--- Table""" & TableName & """" & vbNewLine
    Str = Str & "CREATE TABLE " & TableName & " (" & vbNewLine
    Str = Str & fields
    Str = Str & ");" & vbNewLine
    Str = Str & alters & vbNewLine
    
    CreateTable = Str
End Function

Private Function FileWrite(saveName, data)
    Const adTypeText = 2            'Const value to output
    Const adSaveCreateOverWrite = 2 'Const value to output
    Const adWriteLine = 1
    
    Dim mySrm As Object
    Set mySrm = CreateObject("ADODB.Stream")
    With mySrm
        '*** read ADO in UTF-8 to output
        .Open
        .Type = adTypeText
        .Charset = "UTF-8"
        
        'write an object to a file
        .WriteText data, adWriteLine
        .SaveToFile (saveName), adSaveCreateOverWrite

        'close an object
        .Close
    End With
    
    
    'delete an object from memory
    Set mySrm = Nothing

End Function


Sub generateDDL()
Attribute generateDDL.VB_ProcData.VB_Invoke_Func = "g\n14"
    Dim ddlPath As String
    ddlPath = Sheet1.Range("B1").Text
    If Not fso.FileExists(ddlPath) Then
        MsgBox "TABLEíËã`ExcelÇ™å©Ç¬Ç©ÇËÇ‹ÇπÇÒ. " & vbCrLf & "excelpath" & ddlPath, vbExclamation
        End
    End If
    
    Dim tableHeader As tableHeader
    Set tableHeader = New tableHeader
    tableHeader.Init ThisWorkbook.Sheets("config")
    
    Dim sqlStr As String
    sqlStr = ""
    Dim saveName
    Dim saveDir
    
    Dim ddlBook As Workbook
    Set ddlBook = OpenBook(ddlPath)
    Dim dbSheet As Worksheet
    Dim r As Integer
    r = 2
    For Each dbSheet In ddlBook.Sheets
        If dbSheet.Range("AB4") <> "" Then
            Sheet1.Range("O" & r) = dbSheet.Range("AB4").Text
            saveName = dbSheet.Range("AB4").Text
            dbSheet.Activate
            sqlStr = sqlStr & CreateTable(saveName, tableHeader)
            
            r = r + 1
        End If
    Next

    saveDir = fso.BuildPath(ThisWorkbook.Path, "sql")
    If Not fso.FolderExists(saveDir) Then
        MkDir saveDir
    End If
    
    Dim n As Date
    n = Now
    
    saveName = fso.BuildPath(saveDir, "\ddl_" & Format(n, "yyyy-mm-dd-hh-mm-ss") & ".sql")
    
'    Dim sqlStr As String
'    sqlStr = ""
'    Sheets("table list").Select
'    ' Stop painting
'    Application.ScreenUpdating = False
'
'
'    Do
'        ActiveSheet.Next.Activate
'
'        Dim tableHeader As tableHeader
'
'        Set tableHeader = New tableHeader
'        tableHeader.Init Sheets("config")
'
'        sqlStr = sqlStr & CreateTable(saveName, tableHeader)
'
'    Loop While ActiveSheet.name <> Sheets(Sheets.Count).name ' Loop until last worksheets
    
    ' Write to a file
    Call FileWrite(saveName, sqlStr)
    ' Start painting
    Application.ScreenUpdating = True
    MsgBox "DDL export Done!", vbInformation
End Sub

