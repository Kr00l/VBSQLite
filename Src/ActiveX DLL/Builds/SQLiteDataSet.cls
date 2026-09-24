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
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If
Private Const ARR_BUMP As Long = 10
Private Type TCOLS
Cols() As Variant
End Type
Private Type TRECORDS
Rows() As TCOLS
End Type
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function VarDecFromI8 Lib "oleaut32" (ByVal i64In As Currency, ByRef pDecOut As Variant) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VarDecFromI8 Lib "oleaut32" (ByVal i64In As Currency, ByRef pDecOut As Variant) As Long
#End If
Private Const VT_BYREF As Integer = &H4000
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

#If VBA7 Then
Friend Function FInit(ByVal hStmt As LongPtr) As Boolean
#Else
Friend Function FInit(ByVal hStmt As Long) As Boolean
#End If
Dim Result As Long, ArrUBound As Long, Count As Long, iCol As Long
Dim Int64 As Currency
Result = stub_sqlite3_step(hStmt)
ArrUBound = -1
While Result = SQLITE_ROW
    If PropRecordCount(0) > ArrUBound Then
        ArrUBound = PropRecordCount(0) + ARR_BUMP - 1
        ReDim Preserve PropValues.Rows(0 To ArrUBound) As TCOLS
    End If
    With PropValues.Rows(PropRecordCount(0))
    Count = stub_sqlite3_data_count(hStmt)
    ReDim .Cols(0 To (Count - 1)) As Variant
    For iCol = 0 To (Count - 1)
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
    PropColumns.FInit Me, Nothing, hStmt, PropRecordCount(1)
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

Public Function GetRows(Optional ByVal Rows As Long = -1, Optional ByVal Cols As Long = -1, Optional ByVal Row As Long = -1, Optional ByVal Col As Long = -1, Optional ByVal DataCols As Variant) As Variant
Attribute GetRows.VB_Description = "Copies the content into an two-dimensional variant array."
If Rows < -1 Then Err.Raise 380
If Cols < -1 Then Err.Raise 380
If Row < -1 Then Err.Raise 380
If Col < -1 Then Err.Raise 380
If Row = -1 Then
    Row = Me.Position - 1
    If Row < 0 Then Err.Raise Number:=vbObjectError, Description:="Either BOF or EOF is True"
Else
    If (Row < 0 Or Row > (PropRecordCount(0) - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
End If
If Rows = -1 Then Rows = PropRecordCount(0) - Row
If Rows > 0 Then
    If Rows > (PropRecordCount(0) - Row) Then Rows = PropRecordCount(0) - Row
    Dim LBoundDataCols As Long, UBoundDataCols As Long
    UBoundDataCols = -1
    If Not IsMissing(DataCols) Then
        Dim DataColIndices() As Long, i As Long, j As Long
        If IsArray(DataCols) Then
            Dim Ptr As LongPtr
            CopyMemory Ptr, ByVal UnsignedAdd(VarPtr(DataCols), 8), PTR_SIZE
            Dim VT As Integer
            CopyMemory VT, ByVal VarPtr(DataCols), 2
            If (VT And VT_BYREF) = VT_BYREF Then CopyMemory Ptr, ByVal Ptr, PTR_SIZE
            If Ptr <> NULL_PTR Then
                Dim DimensionCountDataCols As Integer
                CopyMemory DimensionCountDataCols, ByVal Ptr, 2
                If DimensionCountDataCols = 1 Then
                    LBoundDataCols = LBound(DataCols)
                    UBoundDataCols = UBound(DataCols)
                Else
                    Err.Raise Number:=5, Description:="Array must be single dimensioned"
                End If
            Else
                Err.Raise Number:=91, Description:="Array is not allocated"
            End If
            ReDim DataColIndices(LBoundDataCols To UBoundDataCols) As Long
            For i = LBoundDataCols To UBoundDataCols
                If VarType(DataCols(i)) = vbString Then
                    For j = 0 To (PropColumns.Count - 1)
                        If StrComp(PropColumns(j + 1).Name, DataCols(i), vbTextCompare) = 0 Then DataColIndices(i) = j: Exit For
                    Next j
                    If j > (PropColumns.Count - 1) Then Err.Raise 381
                Else
                    DataColIndices(i) = DataCols(i)
                    If DataColIndices(i) < 0 Or DataColIndices(i) > (PropColumns.Count - 1) Then Err.Raise 381
                End If
            Next i
        Else
            LBoundDataCols = 0
            UBoundDataCols = 0
            ReDim DataColIndices(0) As Long
            If VarType(DataCols) = vbString Then
                For j = 0 To (PropColumns.Count - 1)
                    If StrComp(PropColumns(j + 1).Name, DataCols, vbTextCompare) = 0 Then DataColIndices(0) = j: Exit For
                Next j
                If j > (PropColumns.Count - 1) Then Err.Raise 381
            Else
                DataColIndices(0) = DataCols
                If DataColIndices(0) < 0 Or DataColIndices(0) > (PropColumns.Count - 1) Then Err.Raise 381
            End If
        End If
    End If
    Dim ArrRows() As Variant, iRow As Long, iCol As Long
    If (UBoundDataCols - LBoundDataCols) > -1 Then
        If Cols <> -1 Or Col <> -1 Then Err.Raise Number:=5, Description:="Cols and Col must be -1 when DataCols is supplied"
        ReDim ArrRows(0 To (UBoundDataCols - LBoundDataCols), 0 To (Rows - 1)) As Variant
        For iCol = 0 To (UBoundDataCols - LBoundDataCols): For iRow = 0 To (Rows - 1)
            ArrRows(iCol, iRow) = PropValues.Rows(iRow + Row).Cols(DataColIndices(LBoundDataCols + iCol))
        Next iRow, iCol
        GetRows = ArrRows
    Else
        If Col = -1 Then Col = 0
        If (Col < 0 Or Col > (PropRecordCount(1) - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
        If Cols = -1 Then Cols = PropRecordCount(1) - Col
        If Cols > 0 Then
            If Cols > (PropRecordCount(1) - Col) Then Cols = PropRecordCount(1) - Col
            ReDim ArrRows(0 To (Cols - 1), 0 To (Rows - 1)) As Variant
            For iCol = 0 To (Cols - 1): For iRow = 0 To (Rows - 1)
                ArrRows(iCol, iRow) = PropValues.Rows(iRow + Row).Cols(iCol + Col)
            Next iRow, iCol
            GetRows = ArrRows
        End If
    End If
End If
End Function
