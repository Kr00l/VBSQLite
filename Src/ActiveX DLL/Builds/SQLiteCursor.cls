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
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function VarI8FromDec Lib "oleaut32" (ByRef iDecIn As Variant, ByRef i64Out As Currency) As Long
Private Declare PtrSafe Function VarDecFromI8 Lib "oleaut32" (ByVal i64In As Currency, ByRef pDecOut As Variant) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VarI8FromDec Lib "oleaut32" (ByRef iDecIn As Variant, ByRef i64Out As Currency) As Long
Private Declare Function VarDecFromI8 Lib "oleaut32" (ByVal i64In As Currency, ByRef pDecOut As Variant) As Long
#End If
Private Const VT_BYREF As Integer = &H4000
Private Const VT_I8 As Integer = &H14
Private PropHandle As LongPtr
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

#If VBA7 Then
Friend Sub FInit(ByVal hStmt As LongPtr, ByVal DBConnection As SQLiteConnection)
#Else
Friend Sub FInit(ByVal hStmt As Long, ByVal DBConnection As SQLiteConnection)
#End If
PropHandle = hStmt
Set PropDBConnection = DBConnection
PropColumns.FInit Nothing, Me, PropHandle, PropRecordCount(1)
End Sub

Public Property Get Object() As Object
Attribute Object.VB_Description = "Returns the own instance."
Set Object = Me
End Property

#If VBA7 Then
Public Property Get hStmt() As LongPtr
Attribute hStmt.VB_Description = "Returns a SQLite statement handle."
#Else
Public Property Get hStmt() As Long
Attribute hStmt.VB_Description = "Returns a SQLite statement handle."
#End If
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

Public Sub ClearParameters()
Attribute ClearParameters.VB_Description = "Clears all parameter values to null."
If stub_sqlite3_stmt_status(PropHandle, SQLITE_STMTSTATUS_RUN, 1) > 0 Then
    PropLastResult = stub_sqlite3_reset(PropHandle)
    PropRecordCount(0) = 0
    PropPosition = 0
End If
stub_sqlite3_clear_bindings PropHandle
End Sub

Public Function GetParameterCount() As Long
Attribute GetParameterCount.VB_Description = "Returns the index of the largest (rightmost) parameter."
GetParameterCount = stub_sqlite3_bind_parameter_count(PropHandle)
End Function

Public Function GetParameterName(ByVal Index As Long) As String
Attribute GetParameterName.VB_Description = "Returns the name of a parameter given its index."
GetParameterName = SQLiteUTF8PtrToStr(stub_sqlite3_bind_parameter_name(PropHandle, Index))
End Function

Public Property Get ParameterIndex(ByVal Name As String) As Long
Attribute ParameterIndex.VB_Description = "Returns the index of a parameter given its name."
Dim TextUTF8() As Byte
TextUTF8() = UTF16_To_UTF8(Name & vbNullChar)
ParameterIndex = stub_sqlite3_bind_parameter_index(PropHandle, VarPtr(TextUTF8(0)))
End Property

Public Sub SetParameterValue(ByVal Index As Long, ByRef Value As Variant)
Attribute SetParameterValue.VB_Description = "Binds a parameter value to the SQL query."
Dim VT As Integer
VT = VarType(Value)
Dim Result As Long
Select Case VT
    Case vbDecimal, VT_I8, vbLong, vbInteger, vbByte ' SQLITE_INTEGER
        If stub_sqlite3_stmt_status(PropHandle, SQLITE_STMTSTATUS_RUN, 1) > 0 Then
            PropLastResult = stub_sqlite3_reset(PropHandle)
            PropRecordCount(0) = 0
            PropPosition = 0
        End If
        Dim Int64 As Currency
        Select Case VT
            Case vbDecimal
                VarI8FromDec Value, Int64
                Result = stub_sqlite3_bind_int64(PropHandle, Index, Int64)
            Case VT_I8
                CopyMemory Int64, ByVal UnsignedAdd(VarPtr(Value), 8), 8
                Result = stub_sqlite3_bind_int64(PropHandle, Index, Int64)
            Case vbLong, vbInteger, vbByte
                Result = stub_sqlite3_bind_int(PropHandle, Index, Value)
        End Select
    Case vbDouble ' SQLITE_FLOAT
        If stub_sqlite3_stmt_status(PropHandle, SQLITE_STMTSTATUS_RUN, 1) > 0 Then
            PropLastResult = stub_sqlite3_reset(PropHandle)
            PropRecordCount(0) = 0
            PropPosition = 0
        End If
        Result = stub_sqlite3_bind_double(PropHandle, Index, Value)
    Case vbString ' SQLITE_TEXT
        If stub_sqlite3_stmt_status(PropHandle, SQLITE_STMTSTATUS_RUN, 1) > 0 Then
            PropLastResult = stub_sqlite3_reset(PropHandle)
            PropRecordCount(0) = 0
            PropPosition = 0
        End If
        Result = stub_sqlite3_bind_text16(PropHandle, Index, StrPtr(Value), -1, SQLITE_TRANSIENT)
    Case vbArray + vbByte ' SQLITE_BLOB
        If stub_sqlite3_stmt_status(PropHandle, SQLITE_STMTSTATUS_RUN, 1) > 0 Then
            PropLastResult = stub_sqlite3_reset(PropHandle)
            PropRecordCount(0) = 0
            PropPosition = 0
        End If
        Dim pArray As LongPtr, pvData As LongPtr
        CopyMemory pArray, ByVal UnsignedAdd(VarPtr(Value), 8), PTR_SIZE
        CopyMemory VT, ByVal VarPtr(Value), 2
        If (VT And VT_BYREF) = VT_BYREF Then CopyMemory pArray, ByVal pArray, PTR_SIZE
        If pArray <> NULL_PTR Then CopyMemory pvData, ByVal UnsignedAdd(pArray, 12), PTR_SIZE
        If pvData <> NULL_PTR Then
            Result = stub_sqlite3_bind_blob(PropHandle, Index, pvData, LenB(Value), SQLITE_TRANSIENT)
        Else
            Result = stub_sqlite3_bind_null(PropHandle, Index)
        End If
    Case vbNull ' SQLITE_NULL
        If stub_sqlite3_stmt_status(PropHandle, SQLITE_STMTSTATUS_RUN, 1) > 0 Then
            PropLastResult = stub_sqlite3_reset(PropHandle)
            PropRecordCount(0) = 0
            PropPosition = 0
        End If
        Result = stub_sqlite3_bind_null(PropHandle, Index)
    Case Else
        Err.Raise 380
End Select
If Result <> SQLITE_OK Then Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(stub_sqlite3_db_handle(PropHandle)), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(stub_sqlite3_db_handle(PropHandle)))
End Sub

Public Property Get Columns() As SQLiteColumns
Attribute Columns.VB_Description = "Returns a reference to a collection of the column objects."
Attribute Columns.VB_UserMemId = 0
Set Columns = PropColumns
End Property

Public Property Get RecordCount() As Long
Attribute RecordCount.VB_Description = "Returns the number of records accessed in a SQLiteCursor object."
If stub_sqlite3_stmt_status(PropHandle, SQLITE_STMTSTATUS_RUN, 0) = 0 Then Err.Raise Number:=5, Description:="The Cursor has not been initialized"
RecordCount = PropRecordCount(0)
End Property

Public Property Get Position() As Long
Attribute Position.VB_Description = "Returns the position of the current record."
If stub_sqlite3_stmt_status(PropHandle, SQLITE_STMTSTATUS_RUN, 0) = 0 Then Err.Raise Number:=5, Description:="The Cursor has not been initialized"
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
If stub_sqlite3_stmt_status(PropHandle, SQLITE_STMTSTATUS_RUN, 0) = 0 Then Err.Raise Number:=5, Description:="The Cursor has not been initialized"
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
            Value = stub_sqlite3_column_double(PropHandle, Col)
        Case SQLITE_TEXT
            Value = SQLiteUTF8PtrToStr(stub_sqlite3_column_text(PropHandle, Col), stub_sqlite3_column_bytes(PropHandle, Col))
        Case SQLITE_BLOB
            Value = SQLiteBlobToByteArray(stub_sqlite3_column_blob(PropHandle, Col), stub_sqlite3_column_bytes(PropHandle, Col))
        Case SQLITE_NULL
            Value = Null
    End Select
Else
    Err.Raise Number:=vbObjectError, Description:="Either BOF or EOF is True"
End If
End Property

Public Property Get BOF() As Boolean
Attribute BOF.VB_Description = "Returns a value that indicates if the position is before the first record."
If stub_sqlite3_stmt_status(PropHandle, SQLITE_STMTSTATUS_RUN, 0) = 0 Then Err.Raise Number:=5, Description:="The Cursor has not been initialized"
If PropRecordCount(0) = 0 Or PropPosition < 1 Then BOF = True
End Property

Public Property Get EOF() As Boolean
Attribute EOF.VB_Description = "Returns a value that indicates if the position is after the last record."
If stub_sqlite3_stmt_status(PropHandle, SQLITE_STMTSTATUS_RUN, 0) = 0 Then Err.Raise Number:=5, Description:="The Cursor has not been initialized"
If PropRecordCount(0) = 0 Or PropPosition > PropRecordCount(0) Then EOF = True
End Property

Public Sub MoveFirst()
Attribute MoveFirst.VB_Description = "Moves the position to the first record."
If stub_sqlite3_stmt_status(PropHandle, SQLITE_STMTSTATUS_RUN, 1) > 0 Then
    PropLastResult = stub_sqlite3_reset(PropHandle)
    PropRecordCount(0) = 0
    PropPosition = 0
End If
PropLastResult = stub_sqlite3_step(PropHandle)
If PropLastResult = SQLITE_ROW Then
    PropRecordCount(0) = 1
    PropPosition = 1
ElseIf PropLastResult <> SQLITE_DONE Then
    Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(stub_sqlite3_db_handle(PropHandle)), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(stub_sqlite3_db_handle(PropHandle)))
End If
End Sub

Public Sub MoveNext()
Attribute MoveNext.VB_Description = "Moves the position to the next record."
If stub_sqlite3_stmt_status(PropHandle, SQLITE_STMTSTATUS_RUN, 0) = 0 Then Err.Raise Number:=5, Description:="The Cursor has not been initialized"
If PropLastResult = SQLITE_ROW Then
    PropLastResult = stub_sqlite3_step(PropHandle)
    Select Case PropLastResult
        Case SQLITE_ROW
            PropPosition = PropPosition + 1
            PropRecordCount(0) = PropPosition
        Case SQLITE_DONE
            PropPosition = PropPosition + 1
        Case Else
            Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(stub_sqlite3_db_handle(PropHandle)), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(stub_sqlite3_db_handle(PropHandle)))
    End Select
Else
    Err.Raise Number:=vbObjectError, Description:="The Cursor has already reached EOF"
End If
End Sub
