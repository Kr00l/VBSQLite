VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLiteCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function VarDecFromI8 Lib "oleaut32" (ByVal i64In As Currency, ByRef pDecOut As Variant) As Long
Private PropHandle As Long
Private PropDBConnection As SQLiteConnection
Private PropColumns As SQLiteColumns
Private PropLastResult As Long
Private PropRecordCount(0 To 1) As Long
Private PropPosition As Long

Private Sub Class_Initialize()
Set PropColumns = New SQLiteColumns
End Sub

Private Sub Class_Terminate()
stub_sqlite3_finalize PropHandle
Set PropDBConnection = Nothing
Set PropColumns = Nothing
End Sub

Friend Sub FInit(ByVal hStmt As Long, ByVal DBConnection As SQLiteConnection)
PropHandle = hStmt
Set PropDBConnection = DBConnection
PropColumns.FInit Nothing, Me, hStmt, PropRecordCount(1)
PropLastResult = stub_sqlite3_step(hStmt)
If PropLastResult = SQLITE_ROW Then
    PropRecordCount(0) = 1
    PropPosition = 1
End If
End Sub

Public Property Get Object() As Object
Attribute Object.VB_Description = "Returns the own instance."
Set Object = Me
End Property

Public Property Get hStmt() As Long
Attribute hStmt.VB_Description = "Returns a SQLite statement handle."
hStmt = PropHandle
End Property

Public Property Get DBConnection() As SQLiteConnection
Attribute DBConnection.VB_Description = "Returns a reference to a SQLiteConnection object to which the specified SQLiteCursor object belongs to."
Set DBConnection = PropDBConnection
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the text used to create the SQLiteCursor object."
SQL = SQLiteUTF8PtrToStr(stub_sqlite3_sql(PropHandle))
End Property

Public Property Get Columns() As SQLiteColumns
Attribute Columns.VB_Description = "Returns a reference to a collection of the column objects."
Attribute Columns.VB_UserMemId = 0
Set Columns = PropColumns
End Property

Public Property Get RecordCount() As Long
Attribute RecordCount.VB_Description = "Returns the number of records accessed in a SQLiteCursor object."
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

Public Property Get Value(ByVal Col As Long) As Variant
Attribute Value.VB_Description = "Returns the value contents of the current record."
If (Col < 0 Or Col > (PropRecordCount(1) - 1)) Then Err.Raise Number:=381, Description:="Subscript out of range"
If PropLastResult = SQLITE_ROW Then
    Select Case stub_sqlite3_column_type(PropHandle, Col)
        Case SQLITE_INTEGER
            Dim Int64 As Currency
            Int64 = stub_sqlite3_column_int64(PropHandle, Col)
            If Int64 >= -214748.3648@ And Int64 <= 214748.3647@ Then
                Value = stub_sqlite3_column_int(PropHandle, Col)
            Else
                Value = CDec(0)
                VarDecFromI8 Int64, Value
            End If
        Case SQLITE_FLOAT
            Value = stub_sqlite3_column_double(hStmt, Col)
        Case SQLITE_TEXT
            Value = SQLiteUTF8PtrToStr(stub_sqlite3_column_text(hStmt, Col), stub_sqlite3_column_bytes(hStmt, Col))
        Case SQLITE_BLOB
            Value = SQLiteBlobToByteArray(stub_sqlite3_column_blob(hStmt, Col), stub_sqlite3_column_bytes(hStmt, Col))
        Case SQLITE_NULL
            Value = Null
    End Select
Else
    Err.Raise Number:=vbObjectError, Description:="Either BOF or EOF is True"
End If
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
If stub_sqlite3_reset(PropHandle) = SQLITE_OK Then
    PropLastResult = stub_sqlite3_step(PropHandle)
    If PropLastResult = SQLITE_ROW Then PropPosition = 1
Else
    Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(stub_sqlite3_db_handle(PropHandle)), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(stub_sqlite3_db_handle(PropHandle)))
End If
End Sub

Public Sub MoveNext()
Attribute MoveNext.VB_Description = "Moves the position to the next record."
If PropLastResult = SQLITE_ROW Then
    PropLastResult = stub_sqlite3_step(PropHandle)
    Select Case PropLastResult
        Case SQLITE_ROW
            PropPosition = PropPosition + 1
            If PropPosition > PropRecordCount(0) Then PropRecordCount(0) = PropPosition
        Case SQLITE_DONE
            PropPosition = PropPosition + 1
        Case Else
            Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(stub_sqlite3_db_handle(PropHandle)), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(stub_sqlite3_db_handle(PropHandle)))
    End Select
Else
    Err.Raise Number:=vbObjectError, Description:="The Cursor has already reached EOF"
End If
End Sub
