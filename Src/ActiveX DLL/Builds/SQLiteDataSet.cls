VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLiteDataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const ARR_BUMP As Long = 10
Private Type TCOLS
Cols() As Variant
End Type
Private Type TRECORDS
Rows() As TCOLS
End Type
Private Declare Function VarDecFromI8 Lib "oleaut32" (ByVal i64In As Currency, ByRef pDecOut As Variant) As Long
Private PropColumns As SQLiteColumns
Private PropValues As TRECORDS
Private PropRecordCount(0 To 1) As Long
Private PropPosition As Long

Private Sub Class_Initialize()
Set PropColumns = New SQLiteColumns
End Sub

Private Sub Class_Terminate()
Set PropColumns = Nothing
End Sub

Friend Function FInit(ByVal hStmt As Long) As Boolean
Dim Result As Long, ArrUBound As Long, Cols As Long, iCol As Long
Dim Int64 As Currency
Result = stub_sqlite3_step(hStmt)
ArrUBound = -1
While Result = SQLITE_ROW
    If PropRecordCount(0) > ArrUBound Then
        ArrUBound = PropRecordCount(0) + ARR_BUMP - 1
        ReDim Preserve PropValues.Rows(0 To ArrUBound) As TCOLS
    End If
    With PropValues.Rows(PropRecordCount(0))
    Cols = stub_sqlite3_data_count(hStmt)
    ReDim .Cols(0 To (Cols - 1)) As Variant
    For iCol = 0 To (Cols - 1)
        Select Case stub_sqlite3_column_type(hStmt, iCol)
            Case SQLITE_INTEGER
                Int64 = stub_sqlite3_column_int64(hStmt, iCol)
                If Int64 >= -214748.3648@ And Int64 <= 214748.3647@ Then
                    .Cols(iCol) = stub_sqlite3_column_int(hStmt, iCol)
                Else
                    .Cols(iCol) = CDec(0)
                    VarDecFromI8 Int64, .Cols(iCol)
                End If
            Case SQLITE_FLOAT
                .Cols(iCol) = stub_sqlite3_column_double(hStmt, iCol)
            Case SQLITE_TEXT
                .Cols(iCol) = SQLiteUTF8PtrToStr(stub_sqlite3_column_text(hStmt, iCol), stub_sqlite3_column_bytes(hStmt, iCol))
            Case SQLITE_BLOB
                .Cols(iCol) = SQLiteBlobToByteArray(stub_sqlite3_column_blob(hStmt, iCol), stub_sqlite3_column_bytes(hStmt, iCol))
            Case SQLITE_NULL
                .Cols(iCol) = Null
        End Select
    Next iCol
    End With
    PropRecordCount(0) = PropRecordCount(0) + 1
    Result = stub_sqlite3_step(hStmt)
Wend
If Result = SQLITE_DONE Then
    PropColumns.FInit Me, hStmt, PropRecordCount(1)
    If PropRecordCount(0) > 0 Then
        If ArrUBound > (PropRecordCount(0) - 1) Then ReDim Preserve PropValues.Rows(0 To (PropRecordCount(0) - 1)) As TCOLS
        PropPosition = 1
    End If
    FInit = True
End If
End Function

Public Property Get Object() As Object
Attribute Object.VB_Description = "Returns the own instance."
Set Object = Me
End Property

Public Property Get Columns() As SQLiteColumns
Attribute Columns.VB_Description = "Returns a reference to a collection of the column objects."
Attribute Columns.VB_UserMemId = 0
Set Columns = PropColumns
End Property

Public Property Get RecordCount() As Long
Attribute RecordCount.VB_Description = "Returns the number of records accessed in a SQLiteDataSet object."
RecordCount = PropRecordCount(0)
End Property

Public Property Get Position() As Long
Attribute Position.VB_Description = "Returns the position of the current record."
If PropRecordCount(0) > 0 Then
    If PropPosition > PropRecordCount(0) Then
        Position = -3
    ElseIf PropPosition < 1 Then
        Position = -2
    Else
        Position = PropPosition
    End If
Else
    Position = -1
End If
End Property

Public Property Get ValueMatrix(ByVal Row As Long, ByVal Col As Long) As Variant
Attribute ValueMatrix.VB_Description = "Returns the value contents of an arbitrary record (row/col subscripts)."
If (Row < 0 Or Row > (PropRecordCount(0) - 1)) Or (Col < 0 Or Col > (PropRecordCount(1) - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
ValueMatrix = PropValues.Rows(Row).Cols(Col)
End Property

Public Property Get BOF() As Boolean
Attribute BOF.VB_Description = "Returns a value that indicates if the position is before the first record."
If PropRecordCount(0) = 0 Or PropPosition < 1 Then BOF = True
End Property

Public Property Get EOF() As Boolean
Attribute EOF.VB_Description = "Returns a value that indicates if the position is after the last record."
If PropRecordCount(0) = 0 Or PropPosition > PropRecordCount(0) Then EOF = True
End Property

Public Sub MoveFirst()
Attribute MoveFirst.VB_Description = "Moves the position to the first record."
If PropRecordCount(0) > 0 Then PropPosition = 1
End Sub

Public Sub MoveLast()
Attribute MoveLast.VB_Description = "Moves the position to the last record."
If PropRecordCount(0) > 0 Then PropPosition = PropRecordCount(0)
End Sub

Public Sub MoveNext()
Attribute MoveNext.VB_Description = "Moves the position to the next record."
If Me.EOF = False Then
    PropPosition = PropPosition + 1
Else
    Err.Raise Number:=vbObjectError, Description:="The DataSet has already reached EOF"
End If
End Sub

Public Sub MovePrevious()
Attribute MovePrevious.VB_Description = "Moves the position to the previous record."
If Me.BOF = False Then
    PropPosition = PropPosition - 1
Else
    Err.Raise Number:=vbObjectError, Description:="The DataSet has already reached BOF"
End If
End Sub

Public Sub Move(ByVal NumRecords As Long)
Attribute Move.VB_Description = "Moves the position of the current record."
If NumRecords > 0 Then
    If Me.EOF = False Then
        PropPosition = PropPosition + NumRecords
        If PropPosition > PropRecordCount(0) Then PropPosition = PropRecordCount(0) + 1
    Else
        Err.Raise Number:=vbObjectError, Description:="The DataSet has already reached EOF"
    End If
ElseIf NumRecords < 0 Then
    If Me.BOF = False Then
        PropPosition = PropPosition + NumRecords
        If PropPosition < 1 Then PropPosition = 0
    Else
        Err.Raise Number:=vbObjectError, Description:="The DataSet has already reached BOF"
    End If
End If
End Sub
