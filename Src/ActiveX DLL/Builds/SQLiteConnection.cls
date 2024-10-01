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
#If False Then
Private SQLiteReadOnly, SQLiteReadWrite, SQLiteReadWriteCreate
#End If
Public Enum SQLiteOpenMode
SQLiteReadOnly = 0
SQLiteReadWrite = 1
SQLiteReadWriteCreate = 2
End Enum
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function VarDecFromI8 Lib "oleaut32" (ByVal i64In As Currency, ByRef pDecOut As Variant) As Long
Private Declare PtrSafe Function VarI8FromDec Lib "oleaut32" (ByRef iDecIn As Variant, ByRef i64Out As Currency) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VarDecFromI8 Lib "oleaut32" (ByVal i64In As Currency, ByRef pDecOut As Variant) As Long
Private Declare Function VarI8FromDec Lib "oleaut32" (ByRef iDecIn As Variant, ByRef i64Out As Currency) As Long
#End If
Private Const VT_I8 As Integer = &H14
Private PropHandle As LongPtr

Private Sub Class_Initialize()
Call SQLiteAddRef
End Sub

Private Sub Class_Terminate()
If PropHandle <> NULL_PTR Then
    stub_sqlite3_close_v2 PropHandle
    PropHandle = NULL_PTR
End If
Call SQLiteRelease
End Sub

Public Property Get Object() As Object
Attribute Object.VB_Description = "Returns the own instance."
Set Object = Me
End Property

#If VBA7 Then
Public Property Get hDB() As LongPtr
Attribute hDB.VB_Description = "Returns a SQLite database connection handle."
#Else
Public Property Get hDB() As Long
Attribute hDB.VB_Description = "Returns a SQLite database connection handle."
#End If
hDB = PropHandle
End Property

Public Property Get Version() As String
Attribute Version.VB_Description = "Returns the version of the SQLite library."
Version = SQLiteUTF8PtrToStr(stub_sqlite3_libversion())
End Property

Public Sub OpenDB(ByVal FileName As String, Optional ByVal Mode As SQLiteOpenMode = SQLiteReadWriteCreate, Optional ByVal SharedCache As Boolean)
Attribute OpenDB.VB_Description = "Opens the SQLite database connection."
Select Case Mode
    Case SQLiteReadOnly, SQLiteReadWrite, SQLiteReadWriteCreate
        If PropHandle <> NULL_PTR Then
            stub_sqlite3_close_v2 PropHandle
            PropHandle = NULL_PTR
        End If
        Dim FileNameUTF8() As Byte, Flags As Long, Result As Long
        FileNameUTF8() = UTF16_To_UTF8(FileName & vbNullChar)
        Flags = SQLITE_OPEN_FULLMUTEX
        Select Case Mode
            Case SQLiteReadOnly
                Flags = Flags Or SQLITE_OPEN_READONLY
            Case SQLiteReadWrite
                Flags = Flags Or SQLITE_OPEN_READWRITE
            Case SQLiteReadWriteCreate
                Flags = Flags Or SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE
        End Select
        If SharedCache = False Then
            Flags = Flags Or SQLITE_OPEN_PRIVATECACHE
        Else
            Flags = Flags Or SQLITE_OPEN_SHAREDCACHE
        End If
        Result = stub_sqlite3_open_v2(VarPtr(FileNameUTF8(0)), PropHandle, Flags, NULL_PTR)
        If Result <> SQLITE_OK Then
            Dim ErrVal As Long, ErrMsg As String
            If PropHandle <> NULL_PTR Then
                ErrVal = stub_sqlite3_errcode(PropHandle)
                ErrMsg = SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(PropHandle))
                stub_sqlite3_close_v2 PropHandle
                PropHandle = NULL_PTR
            End If
            Err.Raise Number:=vbObjectError + ErrVal, Description:=ErrMsg
        Else
            stub_sqlite3_regexp_init PropHandle, NULL_PTR, NULL_PTR
            Call SQLiteOverloadBuiltinFunctions(PropHandle)
        End If
    Case Else
        Err.Raise 380
End Select
End Sub

Public Sub CloseDB()
Attribute CloseDB.VB_Description = "Closes the SQLite database connection."
If PropHandle <> NULL_PTR Then
    stub_sqlite3_close_v2 PropHandle
    PropHandle = NULL_PTR
End If
End Sub

Public Property Get FileName(Optional ByVal DBName As String = "main") As String
Attribute FileName.VB_Description = "Returns the filename of a SQLite database."
If PropHandle = NULL_PTR Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
Dim DBNameUTF8() As Byte, Result As LongPtr
DBNameUTF8() = UTF16_To_UTF8(DBName & vbNullChar)
Result = stub_sqlite3_db_filename(PropHandle, VarPtr(DBNameUTF8(0)))
If Result <> NULL_PTR Then
    FileName = SQLiteUTF8PtrToStr(Result)
Else
    Err.Raise Number:=380, Description:="DB name not found"
End If
End Property

Public Property Get ReadOnly(Optional ByVal DBName As String = "main") As Boolean
Attribute ReadOnly.VB_Description = "Indicates if a SQLite database is read-only."
If PropHandle = NULL_PTR Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
Dim DBNameUTF8() As Byte, Result As Long
DBNameUTF8() = UTF16_To_UTF8(DBName & vbNullChar)
Result = stub_sqlite3_db_readonly(PropHandle, VarPtr(DBNameUTF8(0)))
If Result <> -1 Then
    ReadOnly = CBool(Result = 1)
Else
    Err.Raise Number:=380, Description:="DB name not found"
End If
End Property

Public Property Get AffectedRows() As Long
Attribute AffectedRows.VB_Description = "Returns the number of rows modified, inserted or deleted by the most recently completed SQL statement. The value returned by this function will not change by read-only SQL statements."
If PropHandle = NULL_PTR Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
AffectedRows = stub_sqlite3_changes(PropHandle)
End Property

Public Property Get LastInsertRowID() As Variant
Attribute LastInsertRowID.VB_Description = "Returns/sets the rowid of the most recent successful insert into a rowid table."
If PropHandle = NULL_PTR Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
Dim Int64 As Currency
Int64 = stub_sqlite3_last_insert_rowid(PropHandle)
If Int64 >= -214748.3648@ And Int64 <= 214748.3647@ Then
    LastInsertRowID = CLng(Int64 * 10000@)
Else
    LastInsertRowID = CDec(0)
    VarDecFromI8 Int64, LastInsertRowID
End If
End Property

Public Property Let LastInsertRowID(ByVal Value As Variant)
If PropHandle = NULL_PTR Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
Dim VT As Integer
VT = VarType(Value)
Select Case VT
    Case vbDecimal, VT_I8, vbLong, vbInteger, vbByte
        Dim Int64 As Currency
        Select Case VT
            Case vbDecimal
                VarI8FromDec Value, Int64
            Case VT_I8
                CopyMemory Int64, ByVal UnsignedAdd(VarPtr(Value), 8), 8
            Case vbLong, vbInteger, vbByte
                Int64 = Value / 10000@
        End Select
        stub_sqlite3_set_last_insert_rowid PropHandle, Int64
    Case Else
        Err.Raise 380
End Select
End Property

Public Property Get AutoCommit() As Boolean
Attribute AutoCommit.VB_Description = "Returns a value that determines whether or not the SQLite database connection is in autocommit mode."
If PropHandle = NULL_PTR Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
AutoCommit = CBool(stub_sqlite3_get_autocommit(PropHandle) <> 0)
End Property

Public Property Let AutoCommit(ByVal Value As Boolean)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Sub Execute(ByVal SQL As String)
Attribute Execute.VB_Description = "Executes a SQL statement."
If PropHandle = NULL_PTR Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
Dim SQLUTF8() As Byte, Result As Long
SQLUTF8() = UTF16_To_UTF8(SQL & vbNullChar)
Result = stub_sqlite3_exec(PropHandle, VarPtr(SQLUTF8(0)), NULL_PTR, NULL_PTR, NULL_PTR)
If Result <> SQLITE_OK Then Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(PropHandle), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(PropHandle))
End Sub

Public Function OpenDataSet(ByVal SQL As String) As SQLiteDataSet
Attribute OpenDataSet.VB_Description = "Runs a SQL query and returns a reference to a SQLiteDataSet object."
If PropHandle = NULL_PTR Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
Dim hStmt As LongPtr, Result As Long
Result = stub_sqlite3_prepare16_v2(PropHandle, StrPtr(SQL), -1, hStmt, NULL_PTR)
If Result = SQLITE_OK Then
    Dim NewDataSet As New SQLiteDataSet
    If NewDataSet.FInit(hStmt) = True Then
        stub_sqlite3_finalize hStmt
        Set OpenDataSet = NewDataSet
    Else
        stub_sqlite3_finalize hStmt
        Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(PropHandle), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(PropHandle))
    End If
Else
    Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(PropHandle), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(PropHandle))
End If
End Function

Public Function CreateCommand(ByVal SQL As String) As SQLiteCommand
Attribute CreateCommand.VB_Description = "Creates a SQL statement in which parameters can be bound."
If PropHandle = NULL_PTR Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
Dim hStmt As LongPtr, Result As Long
Result = stub_sqlite3_prepare16_v3(PropHandle, StrPtr(SQL), -1, SQLITE_PREPARE_PERSISTENT, hStmt, NULL_PTR)
If Result = SQLITE_OK Then
    Set CreateCommand = New SQLiteCommand
    CreateCommand.FInit hStmt, Me
Else
    Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(PropHandle), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(PropHandle))
End If
End Function

Public Function CreateCursor(ByVal SQL As String) As SQLiteCursor
Attribute CreateCursor.VB_Description = "Creates a SQL query in which data can be retrieved and where the cursor is treated as an iterator."
If PropHandle = NULL_PTR Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
Dim hStmt As LongPtr, Result As Long
Result = stub_sqlite3_prepare16_v3(PropHandle, StrPtr(SQL), -1, SQLITE_PREPARE_PERSISTENT, hStmt, NULL_PTR)
If Result = SQLITE_OK Then
    Set CreateCursor = New SQLiteCursor
    CreateCursor.FInit hStmt, Me
Else
    Err.Raise Number:=vbObjectError + stub_sqlite3_errcode(PropHandle), Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(PropHandle))
End If
End Function

Public Sub BackupDB(ByVal Destination As SQLiteConnection, Optional ByVal DestinationDBName As String = "main", Optional ByVal SourceDBName As String = "main")
Attribute BackupDB.VB_Description = "Backups (copies) a SQLite database between two SQLite database connections."
If PropHandle = NULL_PTR Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
If Destination Is Nothing Then
    Err.Raise Number:=35607, Description:="Required argument is missing"
ElseIf Destination.hDB = 0 Then
    Err.Raise Number:=35607, Description:="Required argument is missing"
ElseIf Destination Is Me Then
    Err.Raise Number:=35700, Description:="Circular object referencing is not allowed"
End If
Dim DestinationDBNameUTF8() As Byte, SourceDBNameUTF8() As Byte, pBackup As LongPtr
DestinationDBNameUTF8() = UTF16_To_UTF8(DestinationDBName & vbNullChar)
SourceDBNameUTF8() = UTF16_To_UTF8(SourceDBName & vbNullChar)
pBackup = stub_sqlite3_backup_init(Destination.hDB, VarPtr(DestinationDBNameUTF8(0)), PropHandle, VarPtr(SourceDBNameUTF8(0)))
If pBackup <> NULL_PTR Then
    stub_sqlite3_backup_step pBackup, -1
    stub_sqlite3_backup_finish pBackup
End If
Dim ErrVal As Long
ErrVal = stub_sqlite3_errcode(Destination.hDB)
If ErrVal <> SQLITE_OK Then Err.Raise Number:=vbObjectError + ErrVal, Description:=SQLiteUTF8PtrToStr(stub_sqlite3_errmsg(Destination.hDB))
End Sub

Public Sub SetProgressHandler(ByVal Handler As ISQLiteProgressHandler, Optional ByVal VMInstructions As Long = 100)
Attribute SetProgressHandler.VB_Description = "Registers/unregisters a progress handler callback to be invoked during long running SQL statements. An example use case is to keep a GUI updated and responsive."
If PropHandle = NULL_PTR Then Err.Raise Number:=5, Description:="DB connection must be opened before it can be used"
If VMInstructions < 0 Then Err.Raise 380
If Handler Is Nothing Then VMInstructions = 0
stub_sqlite3_progress_handler PropHandle, VMInstructions, AddressOf SQLiteProgressHandlerCallback, ObjPtr(Handler)
End Sub
