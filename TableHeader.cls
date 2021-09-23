VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public cellTableName As String
Public rowCommentTbl As String
Public lineNoFirstCol As Integer
Public rowColName As String
Public rowDType As String
Public rowLen As String
Public rowPkey As String
Public rowNotNull As String
Public rowConstr As String
Public rowCommentCol As String
Public ownerName As String
Private Sub Class_Initialize()
    cellTableName = "B1"         'Cell of table name
    rowCommentTbl = "E1"         'row name of comment on a table
    lineNoFirstCol = 4           'First column number of filelds
    rowColName = "A"             'row name of physical column name
    rowDType = "B"               'row name of data type
    rowLen = "C"                 'row name of length
    rowPkey = "D"                'row name of PK which is specified or not
    rowNotNull = "E"             'row name of NN which is specified or not
    rowConstr = "F"              'row name of Constrains(FK,UNIQUE)
    rowCommentCol = "G"          'row name of comment on each column
    ownerName = "postgres"
End Sub

Public Sub Init(thisSHeet As Worksheet)
    If thisSHeet.Range("A2").Text <> "#VarName" Then
        MsgBox "configシート設定ミス。A2<>#VarName!", vbExclamation
        End
    End If
    
    cellTableName = thisSHeet.Range("C3").value        'Cell of table name
    rowCommentTbl = thisSHeet.Range("C4").value          'row name of comment on a table
    lineNoFirstCol = thisSHeet.Range("C5").value            'First column number of filelds
    rowColName = thisSHeet.Range("C6").value            'row name of physical column name
    rowDType = thisSHeet.Range("C7").value               'row name of data type
    rowLen = thisSHeet.Range("C8").value                'row name of length
    rowPkey = thisSHeet.Range("C9").value              'row name of PK which is specified or not
    rowNotNull = thisSHeet.Range("C10").value             'row name of NN which is specified or not
    rowConstr = thisSHeet.Range("C11").value             'row name of Constrains(FK,UNIQUE)
    rowCommentCol = thisSHeet.Range("C12").value          'row name of comment on each column
    ownerName = thisSHeet.Range("C13").value
    
End Sub
