VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLiteConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
#If False Then
Private SQLiteReadOnly, SQLiteReadWrite, SQLiteReadWriteCreate
#End If
Public Enum SQLiteOpenMode
SQLiteReadOnly = 0
SQLiteReadWrite = 1
SQLiteReadWriteCreate = 2
End Enum
Private PropHandle As Long

Private Sub Class_Initialize()
Call SQLiteAddRef
End Sub

Private Sub Class_Terminate()
If PropHandle <> 0 Then
    stub_sqlite3_close_v2 PropHandle
    PropHandle = 0
End If
Call SQLiteRelease
End Sub

Public Property Get Object() As Object
Attribute Object.VB_Description = "Returns the own instance."
Set Object = Me
End Property

Public Property Get hDB() As Long
Attribute hDB.VB_Description = "Returns a SQLite database connection handle."
hDB = PropHandle
End Property

Public Property Get Version() As String
Attribute Version.VB_Description = "Returns the version of the SQLite library."
Version = SQLiteUTF8PtrToStr(stub_sqlite3_libversion())
End Property

Public Sub OpenDB(ByVal FileName As String, Optional ByVal Mode As SQLiteOpenMode = SQLiteReadWriteCreate)
Attribute OpenDB.VB_Description = "Opens the SQLite database connection."
Select Case Mode
    Case SQLiteReadOnly, SQLiteReadWrite, SQLiteReadWriteCreate
        If PropHandle <> 0 Then
            stub_sqlite3_close_v2 PropHandle
            PropHandle = 0
        End If
        Dim FileNameUTF8() As Byte, Flags As Long, Result As Long
        FileNameUTF8() = UTF16_To_UTF8(FileName & vbNullChar)
        Select Case Mode
            Case SQLiteReadOnly
                Flags = SQLITE_OPEN_READONLY
            Case SQLiteReadWrite
                Flags = SQLITE_OPEN_READWRITE
            Case SQLiteReadWriteCreate
                Flags = SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE
        End Select
        Result = stub_sqlite3_open_v2(VarPtr(FileNameUTF8(0)), PropHandle, Flags, 0)
        If Result <> SQLITE_OK Then
            Dim ErrVal As Long, ErrMsg As String
            If PropHandle <> 0 Then
                ErrVal = stub_sqlite3_errcode(PropHandle)
                ErrMsg = SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(PropHandle))
                stub_sqlite3_close_v2 PropHandle
                PropHandle = 0
            End If
            Err.Raise Number:=vbObjectError + ErrVal, Description:=ErrMsg
        Else
            Call SQLiteOverloadBuiltinFunctions(PropHandle)
        End If
    Case Else
        Err.Raise 380
End Select
End Sub

Public Sub CloseDB()
Attribute CloseDB.VB_Description = "Closes the SQLite database connection."
If PropHandle <> 0 Then
    stub_sqlite3_close_v2 PropHandle
    PropHandle = 0
End If
End Sub

Public Property Get FileName(Optional ByVal DBName As String = "main") As String
Attribute FileName.VB_Description = "Returns the filename of a SQLite database."
If PropHandle = 0 Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
Dim SQLUTF8() As Byte, Result As Long
SQLUTF8() = UTF16_To_UTF8(DBName & vbNullChar)
Result = stub_sqlite3_db_filename(PropHandle, VarPtr(SQLUTF8(0)))
If Result <> 0 Then
    FileName = SQLiteUTF8PtrToStr(Result)
Else
    Err.Raise Number:=380, Description:="DB name not found"
End If
End Property

Public Property Get ReadOnly(Optional ByVal DBName As String = "main") As Boolean
Attribute ReadOnly.VB_Description = "Indicates if a SQLite database is read-only."
If PropHandle = 0 Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
Dim SQLUTF8() As Byte, Result As Long
SQLUTF8() = UTF16_To_UTF8(DBName & vbNullChar)
Result = stub_sqlite3_db_readonly(PropHandle, VarPtr(SQLUTF8(0)))
If Result <> -1 Then
    ReadOnly = CBool(Result = 1)
Else
    Err.Raise Number:=380, Description:="DB name not found"
End If
End Property

Public Property Get AffectedRows() As Long
Attribute AffectedRows.VB_Description = "Returns the number of rows modified, inserted or deleted by the most recently completed SQL statement. The value returned by this function will not change by read-only SQL statements."
If PropHandle = 0 Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
AffectedRows = stub_sqlite3_changes(PropHandle)
End Property

Public Sub Execute(ByVal SQL As String)
Attribute Execute.VB_Description = "Executes a SQL statement."
If PropHandle = 0 Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
Dim SQLUTF8() As Byte, hStmt As Long, Result As Long
SQLUTF8() = UTF16_To_UTF8(SQL & vbNullChar)
Result = stub_sqlite3_exec(PropHandle, VarPtr(SQLUTF8(0)), 0, 0, 0)
If Result <> SQLITE_OK Then Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(PropHandle), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(PropHandle))
End Sub

Public Function OpenDataSet(ByVal SQL As String) As SQLiteDataSet
Attribute OpenDataSet.VB_Description = "Runs a SQL query and returns a reference to a SQLiteDataSet object."
If PropHandle = 0 Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
Dim SQLUTF8() As Byte, hStmt As Long, Result As Long
SQLUTF8() = UTF16_To_UTF8(SQL & vbNullChar)
Result = stub_sqlite3_prepare_v2(PropHandle, VarPtr(SQLUTF8(0)), UBound(SQLUTF8()) + 1, hStmt, 0)
If Result = SQLITE_OK Then
    If stub_sqlite3_stmt_readonly(hStmt) <> 0 Then
        Dim NewDataSet As New SQLiteDataSet
        If NewDataSet.FInit(hStmt) = True Then
            stub_sqlite3_finalize hStmt
            Set OpenDataSet = NewDataSet
        Else
            stub_sqlite3_finalize hStmt
            Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(PropHandle), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(PropHandle))
        End If
    Else
        stub_sqlite3_finalize hStmt
        Err.Raise Number:=5, Description:="SQL statement must be read-only"
    End If
Else
    Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(PropHandle), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(PropHandle))
End If
End Function

Public Function CreateCommand(ByVal SQL As String) As SQLiteCommand
Attribute CreateCommand.VB_Description = "Creates a SQL statement in which parameters can be bound."
If PropHandle = 0 Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
Dim SQLUTF8() As Byte, hStmt As Long, Result As Long
SQLUTF8() = UTF16_To_UTF8(SQL & vbNullChar)
Result = stub_sqlite3_prepare_v3(PropHandle, VarPtr(SQLUTF8(0)), UBound(SQLUTF8()) + 1, SQLITE_PREPARE_PERSISTENT, hStmt, 0)
If Result = SQLITE_OK Then
    Set CreateCommand = New SQLiteCommand
    CreateCommand.FInit hStmt, Me
Else
    Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(PropHandle), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(PropHandle))
End If
End Function
