Option Explicit

Public cellTableName As String
Public rowCommentTbl As String
Public lineNoFirstCol As Integer
Public rowColName As String
Public rowDType As String
Public rowLen As String 
Public rowPkey As String   
Public rowNotNull As String  
Public rowReferences As String
Public rowReferencingTable As String
Public rowUnique As String
Public rowUniqueColumn As String
Public rowCommentCol As String

Private Sub Class_Initialize()
    cellTableName         = "B1" 'Cell of table name
    rowCommentTbl         = "E1" 'row name of comment on a table
    lineNoFirstCol        = 4    'First column number of filelds
    rowColName            = "A"  'row name of physical column name
    rowDType              = "B"  'row name of data type
    rowLen                = "C"  'row name of length
    rowPkey               = "D"  'row name of PK which is specified or not
    rowNotNull            = "E"  'row name of NN which is specified or not
    rowReferences         = "F"  'row name of FK which is specified or not
    rowReferencingTable   = "G"  'row name of Referencing Table list
    rowUnique             = "H"  'row name of UNIQUE which is specified or not
    rowUniqueColumn       = "I"  'row name of Unique Column
    rowCommentCol         = "J"  'row name of comment on each column
End Sub
