VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLiteCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VarI8FromDec Lib "oleaut32" (ByRef iDecIn As Variant, ByRef i64Out As Currency) As Long
Private Const VT_BYREF As Integer = &H4000
Private Const VT_I8 As Integer = &H14
Private PropHandle As Long
Private PropDBConnection As SQLiteConnection

Private Sub Class_Terminate()
stub_sqlite3_finalize PropHandle
Set PropDBConnection = Nothing
End Sub

Friend Sub FInit(ByVal hStmt As Long, ByVal DBConnection As SQLiteConnection)
PropHandle = hStmt
Set PropDBConnection = DBConnection
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
Attribute DBConnection.VB_Description = "Returns a reference to a SQLiteConnection object to which the specified SQLiteCommand object belongs to."
Set DBConnection = PropDBConnection
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the text used to create the SQLiteCommand object."
SQL = SQLiteUTF8PtrToStr(stub_sqlite3_sql(PropHandle))
End Property

Public Sub ClearParameters()
Attribute ClearParameters.VB_Description = "Clears all parameter values to null."
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
Attribute ParameterIndex.VB_UserMemId = 0
Dim TextUTF8() As Byte
TextUTF8() = UTF16_To_UTF8(Name & vbNullChar)
ParameterIndex = stub_sqlite3_bind_parameter_index(PropHandle, VarPtr(TextUTF8(0)))
End Property

Public Sub SetParameterValue(ByVal Index As Long, ByRef Value As Variant)
Attribute SetParameterValue.VB_Description = "Binds a parameter value to the SQL statement."
Dim VT As Integer
VT = VarType(Value)
Dim Result As Long
Select Case VT
    Case vbDecimal, VT_I8, vbLong, vbInteger, vbByte ' SQLITE_INTEGER
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
        Result = stub_sqlite3_bind_double(PropHandle, Index, Value)
    Case vbString ' SQLITE_TEXT
        Result = stub_sqlite3_bind_text16(PropHandle, Index, StrPtr(Value), -1, SQLITE_TRANSIENT)
    Case vbArray + vbByte ' SQLITE_BLOB
        Dim pArray As Long, pvData As Long
        CopyMemory pArray, ByVal UnsignedAdd(VarPtr(Value), 8), 4
        CopyMemory VT, ByVal VarPtr(Value), 2
        If (VT And VT_BYREF) = VT_BYREF Then CopyMemory pArray, ByVal pArray, 4
        If pArray <> 0 Then CopyMemory pvData, ByVal UnsignedAdd(pArray, 12), 4
        If pvData <> 0 Then
            Result = stub_sqlite3_bind_blob(PropHandle, Index, pvData, LenB(Value), SQLITE_TRANSIENT)
        Else
            Result = stub_sqlite3_bind_null(PropHandle, Index)
        End If
    Case vbNull ' SQLITE_NULL
        Result = stub_sqlite3_bind_null(PropHandle, Index)
    Case Else
        Err.Raise 380
End Select
If Result <> SQLITE_OK Then Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(stub_sqlite3_db_handle(PropHandle)), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(stub_sqlite3_db_handle(PropHandle)))
End Sub

Public Sub Execute()
Attribute Execute.VB_Description = "Executes a SQL statement."
While stub_sqlite3_step(PropHandle) = SQLITE_ROW: Wend
If stub_sqlite3_reset(PropHandle) <> SQLITE_OK Then Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(stub_sqlite3_db_handle(PropHandle)), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(stub_sqlite3_db_handle(PropHandle)))
End Sub

Public Function OpenDataSet() As SQLiteDataSet
Attribute OpenDataSet.VB_Description = "Runs a SQL query and returns a reference to a SQLiteDataSet object."
If stub_sqlite3_stmt_readonly(PropHandle) <> 0 Then
    Dim NewDataSet As New SQLiteDataSet
    If NewDataSet.FInit(PropHandle) = True Then
        stub_sqlite3_reset PropHandle
        Set OpenDataSet = NewDataSet
    Else
        stub_sqlite3_reset PropHandle
        Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(stub_sqlite3_db_handle(PropHandle)), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(stub_sqlite3_db_handle(PropHandle)))
    End If
Else
    Err.Raise Number:=5, Description:="SQL statement must be read-only"
End If
End Function

Public Function CreateCursor() As SQLiteCursor
Attribute CreateCursor.VB_Description = "Creates an expanded SQL query and returns a reference to a SQLiteCursor object."
Dim Ptr As Long
Ptr = stub_sqlite3_expanded_sql(PropHandle)
If Ptr <> 0 Then
    Dim hStmt As Long, Result As Long
    Result = stub_sqlite3_prepare_v3(stub_sqlite3_db_handle(PropHandle), Ptr, -1, SQLITE_PREPARE_PERSISTENT, hStmt, 0)
    stub_sqlite3_free Ptr
    If Result = SQLITE_OK Then
        If stub_sqlite3_stmt_readonly(hStmt) <> 0 Then
            Set CreateCursor = New SQLiteCursor
            If CreateCursor.FInit(hStmt, PropDBConnection) = False Then
                Set CreateCursor = Nothing
                Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(stub_sqlite3_db_handle(PropHandle)), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(stub_sqlite3_db_handle(PropHandle)))
            End If
        Else
            stub_sqlite3_finalize hStmt
            Err.Raise Number:=5, Description:="SQL statement must be read-only"
        End If
    Else
        Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(stub_sqlite3_db_handle(PropHandle)), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(stub_sqlite3_db_handle(PropHandle)))
    End If
Else
    Err.Raise Number:=vbObjectError + SQLITE_NOMEM, Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errstr(SQLITE_NOMEM))
End If
End Function
